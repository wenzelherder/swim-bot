const express = require('express');
const line = require('@line/bot-sdk');
const { google } = require('googleapis');

const app = express();

// ===== 設定區域 =====
const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const SPREADSHEET_ID = process.env.GOOGLE_SPREADSHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || '請假紀錄';

const client = new line.Client(config);

// ===== 課表器材偵測 =====
function detectEquipment(text) {
  const equipment = [];

  const hasFin = /\bFin\b/i.test(text);
  const hasBoard = /\bboard\b/i.test(text);
  const hasPaddle = /\bpaddle\b|\bP\+board\b|\bP \+ board\b/i.test(text);
  const hasPullBuoy = /\bpull\b/i.test(text);

  if (hasFin) equipment.push('🐸 蛙鞋 (Fin)');
  if (hasBoard) equipment.push('🏄 浮板 (Board)');
  if (hasPaddle) equipment.push('🖐️ 划手板 (Paddle)');
  if (hasPullBuoy && !hasPaddle) equipment.push('🟡 浮球 (Pull Buoy)');

  return equipment;
}

function isScheduleMessage(text) {
  // 課表通常包含距離、組數、時間等關鍵字
  return (
    /\d+m\s*x\d+/i.test(text) ||
    /T-\d+:\d+/.test(text) ||
    /w-up|warm.?up|cool.?down|dryland/i.test(text) ||
    /kick|pull|swim|dive/i.test(text)
  );
}

// ===== 請假關鍵字 =====
const LEAVE_KEYWORDS = [
  '請假', '缺席', '不來', '無法出席', '不能來', '沒辦法來',
  '有事', '不能參加', '無法參加', '請個假', '這次不來',
  '今天不來', '明天不來', '這週不來', '下週不來',
];

// 待處理的請假流程（key: userId）
// value: { userName, date, reason, needsConfirm, step }
// step: 'awaiting_date' | 'awaiting_reason' | 'awaiting_confirm'
const pendingConfirmations = new Map();

function nextStep(pending) {
  if (!pending.date) return 'awaiting_date';
  if (!pending.reason) return 'awaiting_reason';
  if (pending.needsConfirm) return 'awaiting_confirm';
  return 'ready';
}

function buildStepMessage(pending) {
  switch (pending.step) {
    case 'awaiting_date':
      return '📅 請問請假日期是哪天？\n（例如：3/15、3月15日、今天、明天）';
    case 'awaiting_reason':
      return '📝 請問請假原因是什麼？\n（輸入原因，或回覆「略過」跳過）';
    case 'awaiting_confirm':
      return `請確認請假資訊是否正確：\n👤 ${pending.userName}\n📅 ${pending.date}\n📝 ${pending.reason === '未說明' ? '原因未說明' : pending.reason}\n\n回覆「是」確認登記，回覆「否」取消。`;
  }
}

/**
 * 回傳：
 *  'none'    - 不是請假（疑問句 or 說別人）
 *  'direct'  - 確定是本人請假（含「我」）
 *  'confirm' - 曖昧，需要二次確認
 */
function classifyLeaveMessage(text) {
  if (!LEAVE_KEYWORDS.some(kw => text.includes(kw))) return 'none';

  // 疑問句過濾：含嗎、呢、？ 或英文問號
  if (/[嗎呢？?]/.test(text)) return 'none';

  const hasFirstPerson = /我|本人/.test(text);
  const hasOtherPerson = /你|妳|他|她|誰|對方/.test(text);

  // 在說別人且非自己
  if (hasOtherPerson && !hasFirstPerson) return 'none';

  // 明確第一人稱
  if (hasFirstPerson) return 'direct';

  // 沒有人稱代詞，曖昧情況
  return 'confirm';
}

function extractDate(text) {
  const now = new Date();
  const mdMatch = text.match(/(\d{1,2})\/(\d{1,2})/);
  if (mdMatch) return `${now.getFullYear()}/${mdMatch[1]}/${mdMatch[2]}`;
  const cnMatch = text.match(/(\d{1,2})月(\d{1,2})日?/);
  if (cnMatch) return `${now.getFullYear()}/${cnMatch[1]}/${cnMatch[2]}`;
  const weekdays = ['日', '一', '二', '三', '四', '五', '六'];
  const weekdayMatch = text.match(/(?:這|下)週?([一二三四五六日])/);
  if (weekdayMatch) {
    const targetDay = weekdays.indexOf(weekdayMatch[1]);
    const isNextWeek = text.includes('下週') || text.includes('下星期');
    const d = new Date(now);
    let diff = targetDay - d.getDay();
    if (isNextWeek) diff += 7;
    else if (diff <= 0) diff += 7;
    d.setDate(d.getDate() + diff);
    return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
  }
  if (text.includes('今天')) return `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()}`;
  if (text.includes('明天')) {
    const d = new Date(now);
    d.setDate(d.getDate() + 1);
    return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
  }
  return '日期未指定';
}

function extractReason(text) {
  const reasonPatterns = [
    /因為(.+?)(?:[，,。\n]|$)/,
    /有(.+?)(?:事|活動|課|比賽)/,
    /去(.+?)(?:[，,。\n]|$)/,
  ];
  for (const pattern of reasonPatterns) {
    const match = text.match(pattern);
    if (match) return match[0].replace(/[，,。]/g, '').trim();
  }
  return '未說明';
}

async function appendToSheet(name, date, reason, timestamp) {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:D`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[name, date, reason, timestamp]],
    },
  });
}

async function ensureSheetHeader() {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const sheets = google.sheets({ version: 'v4', auth });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1`,
    });
    if (!res.data.values || res.data.values.length === 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A1:D1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [['姓名', '請假日期', '請假原因', '登記時間']],
        },
      });
      console.log('已建立標題列');
    }
  } catch (e) {
    console.error('初始化 Sheet 失敗:', e.message);
  }
}

function formatTimestamp(now) {
  return `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()} ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
}

async function getUserName(event) {
  try {
    if (event.source.type === 'group') {
      const profile = await client.getGroupMemberProfile(event.source.groupId, event.source.userId);
      return profile.displayName;
    }
    const profile = await client.getProfile(event.source.userId);
    return profile.displayName;
  } catch (e) {
    console.error('取得用戶名稱失敗:', e.message);
    return '未知成員';
  }
}

async function recordLeave(event, userName, date, reason) {
  const now = new Date();
  const timestamp = formatTimestamp(now);
  await appendToSheet(userName, date, reason, timestamp);
  return client.replyMessage(event.replyToken, {
    type: 'text',
    text: `✅ 已登記請假\n👤 ${userName}\n📅 ${date}\n📝 ${reason === '未說明' ? '原因未說明' : reason}`,
  });
}

async function handleEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') return null;

  const text = event.message.text.trim();
  const senderId = event.source.userId;

  // ===== 處理進行中的請假流程 =====
  if (pendingConfirmations.has(senderId)) {
    const pending = pendingConfirmations.get(senderId);

    // 任何步驟都可以取消
    if (/^(否|不|取消|不要|no)$/i.test(text)) {
      pendingConfirmations.delete(senderId);
      return client.replyMessage(event.replyToken, { type: 'text', text: '已取消，不登記請假。' });
    }

    if (pending.step === 'awaiting_date') {
      const parsedDate = extractDate(text);
      if (parsedDate === '日期未指定') {
        return client.replyMessage(event.replyToken, {
          type: 'text',
          text: '無法辨識日期，請重新輸入。\n（例如：3/15、3月15日、今天、明天）',
        });
      }
      pending.date = parsedDate;
      pending.step = nextStep(pending);
      if (pending.step === 'ready') {
        pendingConfirmations.delete(senderId);
        try { return await recordLeave(event, pending.userName, pending.date, pending.reason); }
        catch (e) { return client.replyMessage(event.replyToken, { type: 'text', text: '⚠️ 登記失敗，請管理員手動記錄。' }); }
      }
      pendingConfirmations.set(senderId, pending);
      return client.replyMessage(event.replyToken, { type: 'text', text: buildStepMessage(pending) });
    }

    if (pending.step === 'awaiting_reason') {
      pending.reason = /^略過$/.test(text) ? '未說明' : text;
      pending.step = nextStep(pending);
      if (pending.step === 'ready') {
        pendingConfirmations.delete(senderId);
        try { return await recordLeave(event, pending.userName, pending.date, pending.reason); }
        catch (e) { return client.replyMessage(event.replyToken, { type: 'text', text: '⚠️ 登記失敗，請管理員手動記錄。' }); }
      }
      pendingConfirmations.set(senderId, pending);
      return client.replyMessage(event.replyToken, { type: 'text', text: buildStepMessage(pending) });
    }

    if (pending.step === 'awaiting_confirm') {
      if (/^(是|對|好|確認|要|yes)$/i.test(text)) {
        pendingConfirmations.delete(senderId);
        try { return await recordLeave(event, pending.userName, pending.date, pending.reason); }
        catch (e) { return client.replyMessage(event.replyToken, { type: 'text', text: '⚠️ 登記失敗，請管理員手動記錄。' }); }
      }
      // 非是/否，再次提示
      return client.replyMessage(event.replyToken, {
        type: 'text',
        text: `請回覆「是」確認登記，或「否」取消。\n\n${buildStepMessage(pending)}`,
      });
    }

    // 未知狀態，清除
    pendingConfirmations.delete(senderId);
  }

  // ===== 課表器材提醒 =====
  if (isScheduleMessage(text)) {
    const equipment = detectEquipment(text);
    if (equipment.length > 0) {
      const replyMsg = `📋 今日課表器材提醒\n\n請準備以下器材：\n${equipment.map(e => `  ${e}`).join('\n')}\n\n祝練習順利！🏊`;
      return client.replyMessage(event.replyToken, { type: 'text', text: replyMsg });
    }
    return null;
  }

  // ===== 請假意圖辨識 =====
  const classification = classifyLeaveMessage(text);
  if (classification === 'none') return null;

  const userName = await getUserName(event);
  const date = extractDate(text);
  const reason = extractReason(text);

  const pending = {
    userName,
    date: date !== '日期未指定' ? date : null,
    reason: reason !== '未說明' ? reason : null,
    needsConfirm: classification === 'confirm',
  };
  pending.step = nextStep(pending);

  // 資料齊全且不需確認 → 直接登記
  if (pending.step === 'ready') {
    try {
      return await recordLeave(event, userName, date, reason);
    } catch (e) {
      console.error('處理失敗:', e.message);
      return client.replyMessage(event.replyToken, { type: 'text', text: '⚠️ 登記失敗，請管理員手動記錄。' });
    }
  }

  // 缺資料或需確認 → 進入多步驟流程
  pendingConfirmations.set(senderId, pending);
  const intro = classification === 'confirm'
    ? `❓ 偵測到請假訊息，${userName} 請確認以下資訊：\n`
    : '';
  return client.replyMessage(event.replyToken, {
    type: 'text',
    text: intro + buildStepMessage(pending),
  });
}

// ===== 路由 =====
app.post('/webhook', line.middleware(config), (req, res) => {
  Promise.all(req.body.events.map(handleEvent))
    .then(() => res.json({ status: 'ok' }))
    .catch(err => {
      console.error(err);
      res.status(500).end();
    });
});

app.get('/', (req, res) => res.send('🏊 游泳隊請假機器人運作中'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
  console.log(`Server running on port ${PORT}`);
  await ensureSheetHeader();
});
