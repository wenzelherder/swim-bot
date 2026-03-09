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

// ===== 練習日 =====
const PRACTICE_DAYS = new Set([2, 4, 6]); // 週二、四、六

function isPracticeDay(dateStr) {
  // dateStr 格式：YYYY/M/D
  const d = new Date(dateStr);
  return PRACTICE_DAYS.has(d.getDay());
}

function weekdayName(dateStr) {
  const names = ['日', '一', '二', '三', '四', '五', '六'];
  return names[new Date(dateStr).getDay()];
}

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

// ===== 取消請假關鍵字 =====
const CANCEL_KEYWORDS = ['取消請假', '取消假', '不請假了', '我不請假了', '取消掉請假'];

function isCancelMessage(text) {
  return CANCEL_KEYWORDS.some(kw => text.includes(kw));
}

// 待處理的請假流程（key: userId）
// type: 'leave' | 'cancel'
// leave steps: 'awaiting_date' | 'awaiting_reason' | 'awaiting_confirm' | 'awaiting_overwrite_confirm'
// cancel steps: 'cancel_awaiting_date' | 'cancel_awaiting_confirm'
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
 *  'none'    - 不是請假（疑問句 or 說別人 or 取消請假）
 *  'direct'  - 確定是本人請假（含「我」）
 *  'confirm' - 曖昧，需要二次確認
 */
function classifyLeaveMessage(text) {
  if (isCancelMessage(text)) return 'none';
  if (!LEAVE_KEYWORDS.some(kw => text.includes(kw))) return 'none';

  // 疑問句過濾：含嗎、呢、？ 或英文問號
  if (/[嗎呢？?]/.test(text)) return 'none';

  const hasFirstPerson = /我|本人/.test(text);
  const hasOtherPerson = /你|妳|他|她|誰|對方/.test(text);

  if (hasOtherPerson && !hasFirstPerson) return 'none';
  if (hasFirstPerson) return 'direct';
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

// ===== Google Sheets 操作 =====
async function getAuth() {
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function appendToSheet(name, date, reason, timestamp) {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:E`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[name, date, reason, timestamp, '']],
    },
  });
}

async function getSheetRows() {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:E`,
  });
  return res.data.values || [];
}

// 查某人本月請假紀錄
async function getMyLeaves(userName) {
  const rows = await getSheetRows();
  const now = new Date();
  const prefix = `${now.getFullYear()}/${now.getMonth() + 1}/`;
  return rows.slice(1).filter(r => r[0] === userName && r[1] && r[1].startsWith(prefix) && r[4] !== '已取消');
}

// 查某天請假名單
async function getLeavesOnDate(date) {
  const rows = await getSheetRows();
  return rows.slice(1).filter(r => r[1] === date && r[4] !== '已取消');
}

// 查本週（週二、四、六）請假名單
async function getWeekLeaves() {
  const rows = await getSheetRows();
  const now = new Date();
  const thisWeek = [];
  // 找出本週的週二、四、六日期
  for (const day of [2, 4, 6]) {
    const d = new Date(now);
    const diff = day - d.getDay();
    d.setDate(d.getDate() + diff);
    const dateStr = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
    thisWeek.push(dateStr);
  }
  return { dates: thisWeek, rows: rows.slice(1).filter(r => thisWeek.includes(r[1]) && r[4] !== '已取消') };
}

// 找到同名同日期且未取消的紀錄，回傳 { rowIndex（1-indexed）, row }
async function findActiveRow(name, date) {
  const rows = await getSheetRows();
  for (let i = 1; i < rows.length; i++) {
    const [rowName, rowDate, , , rowStatus] = rows[i];
    if (rowName === name && rowDate === date && rowStatus !== '已取消') {
      return { rowIndex: i + 1, row: rows[i] };
    }
  }
  return null;
}

async function markRowCancelled(rowIndex) {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!E${rowIndex}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [['已取消']] },
  });
}

async function overwriteRow(rowIndex, name, date, reason, timestamp) {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A${rowIndex}:E${rowIndex}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[name, date, reason, timestamp, '']] },
  });
}

async function ensureSheetHeader() {
  try {
    const auth = await getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1`,
    });
    if (!res.data.values || res.data.values.length === 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A1:E1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [['姓名', '請假日期', '請假原因', '登記時間', '狀態']],
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

// 登記前先檢查重複；有重複則進入覆蓋確認流程
async function checkAndRecord(event, senderId, pending) {
  let found;
  try {
    found = await findActiveRow(pending.userName, pending.date);
  } catch (e) {
    console.error('查詢 Sheet 失敗:', e.message);
    found = null;
  }

  if (found) {
    const existingReason = found.row[2] || '未說明';
    pending.step = 'awaiting_overwrite_confirm';
    pending.existingRowIndex = found.rowIndex;
    pendingConfirmations.set(senderId, pending);
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: `⚠️ 你已有 ${pending.date} 的請假紀錄（原因：${existingReason}）。\n要更新原因為「${pending.reason === '未說明' ? '未說明' : pending.reason}」嗎？\n\n回覆「是」更新，「否」保留原本。`,
    });
  }

  pendingConfirmations.delete(senderId);
  try { return await recordLeave(event, pending.userName, pending.date, pending.reason); }
  catch (e) { return client.replyMessage(event.replyToken, { type: 'text', text: '⚠️ 登記失敗，請管理員手動記錄。' }); }
}

async function handleEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') return null;

  const text = event.message.text.trim();
  const senderId = event.source.userId;

  // ===== 處理進行中的流程 =====
  if (pendingConfirmations.has(senderId)) {
    const pending = pendingConfirmations.get(senderId);

    // 任何步驟都可以中止
    if (/^(否|不|取消|不要|no)$/i.test(text)) {
      pendingConfirmations.delete(senderId);
      const msg = pending.type === 'cancel' ? '好的，保留請假紀錄。' : '已取消，不登記請假。';
      return client.replyMessage(event.replyToken, { type: 'text', text: msg });
    }

    // ===== 取消請假流程 =====
    if (pending.type === 'cancel') {
      if (pending.step === 'cancel_awaiting_date') {
        const date = extractDate(text);
        if (date === '日期未指定') {
          return client.replyMessage(event.replyToken, {
            type: 'text',
            text: '無法辨識日期，請重新輸入。\n（例如：3/15、3月15日、今天、明天）',
          });
        }
        let found;
        try { found = await findActiveRow(pending.userName, date); }
        catch (e) { found = null; }

        if (!found) {
          pendingConfirmations.delete(senderId);
          return client.replyMessage(event.replyToken, { type: 'text', text: `找不到你 ${date} 的請假紀錄。` });
        }
        const existingReason = found.row[2] || '未說明';
        pending.date = date;
        pending.rowIndex = found.rowIndex;
        pending.step = 'cancel_awaiting_confirm';
        pendingConfirmations.set(senderId, pending);
        return client.replyMessage(event.replyToken, {
          type: 'text',
          text: `找到以下請假紀錄：\n👤 ${pending.userName}\n📅 ${date}\n📝 ${existingReason}\n\n確認取消嗎？回覆「是」確認，「否」保留。`,
        });
      }

      if (pending.step === 'cancel_awaiting_confirm') {
        if (/^(是|對|好|確認|要|yes)$/i.test(text)) {
          pendingConfirmations.delete(senderId);
          try {
            await markRowCancelled(pending.rowIndex);
            return client.replyMessage(event.replyToken, { type: 'text', text: `✅ 已取消 ${pending.date} 的請假紀錄。` });
          } catch (e) {
            return client.replyMessage(event.replyToken, { type: 'text', text: '⚠️ 取消失敗，請管理員手動處理。' });
          }
        }
        return client.replyMessage(event.replyToken, {
          type: 'text',
          text: `請回覆「是」確認取消，或「否」保留。\n\n找到以下請假紀錄：\n👤 ${pending.userName}\n📅 ${pending.date}`,
        });
      }

      pendingConfirmations.delete(senderId);
      return null;
    }

    // ===== 請假流程 =====
    if (pending.step === 'awaiting_date') {
      const parsedDate = extractDate(text);
      if (parsedDate === '日期未指定') {
        return client.replyMessage(event.replyToken, {
          type: 'text',
          text: '無法辨識日期，請重新輸入。\n（例如：3/15、3月15日、今天、明天）',
        });
      }
      if (!isPracticeDay(parsedDate)) {
        pendingConfirmations.delete(senderId);
        return client.replyMessage(event.replyToken, {
          type: 'text',
          text: `${parsedDate}（週${weekdayName(parsedDate)}）沒有練習，不需要請假。\n練習日為每週二、四、六。`,
        });
      }
      pending.date = parsedDate;
      pending.step = nextStep(pending);
      if (pending.step === 'ready') return await checkAndRecord(event, senderId, pending);
      pendingConfirmations.set(senderId, pending);
      return client.replyMessage(event.replyToken, { type: 'text', text: buildStepMessage(pending) });
    }

    if (pending.step === 'awaiting_reason') {
      pending.reason = /^略過$/.test(text) ? '未說明' : text;
      pending.step = nextStep(pending);
      if (pending.step === 'ready') return await checkAndRecord(event, senderId, pending);
      pendingConfirmations.set(senderId, pending);
      return client.replyMessage(event.replyToken, { type: 'text', text: buildStepMessage(pending) });
    }

    if (pending.step === 'awaiting_confirm') {
      if (/^(是|對|好|確認|要|yes)$/i.test(text)) {
        return await checkAndRecord(event, senderId, pending);
      }
      return client.replyMessage(event.replyToken, {
        type: 'text',
        text: `請回覆「是」確認登記，或「否」取消。\n\n${buildStepMessage(pending)}`,
      });
    }

    if (pending.step === 'awaiting_overwrite_confirm') {
      if (/^(是|對|好|確認|要|yes)$/i.test(text)) {
        pendingConfirmations.delete(senderId);
        try {
          const now = new Date();
          const timestamp = formatTimestamp(now);
          await overwriteRow(pending.existingRowIndex, pending.userName, pending.date, pending.reason, timestamp);
          return client.replyMessage(event.replyToken, {
            type: 'text',
            text: `✅ 已更新請假紀錄\n👤 ${pending.userName}\n📅 ${pending.date}\n📝 ${pending.reason === '未說明' ? '原因未說明' : pending.reason}`,
          });
        } catch (e) {
          return client.replyMessage(event.replyToken, { type: 'text', text: '⚠️ 更新失敗，請管理員手動處理。' });
        }
      }
      pendingConfirmations.delete(senderId);
      return client.replyMessage(event.replyToken, { type: 'text', text: '好的，保留原本的請假紀錄。' });
    }

    // 未知狀態，清除
    pendingConfirmations.delete(senderId);
  }

  // ===== 取消請假意圖 =====
  if (isCancelMessage(text)) {
    const userName = await getUserName(event);
    const date = extractDate(text);

    if (date === '日期未指定') {
      pendingConfirmations.set(senderId, { type: 'cancel', userName, date: null, step: 'cancel_awaiting_date' });
      return client.replyMessage(event.replyToken, { type: 'text', text: '📅 請問要取消哪一天的請假？' });
    }

    let found;
    try { found = await findActiveRow(userName, date); }
    catch (e) { found = null; }

    if (!found) {
      return client.replyMessage(event.replyToken, { type: 'text', text: `找不到你 ${date} 的請假紀錄。` });
    }
    const existingReason = found.row[2] || '未說明';
    pendingConfirmations.set(senderId, { type: 'cancel', userName, date, rowIndex: found.rowIndex, step: 'cancel_awaiting_confirm' });
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: `找到以下請假紀錄：\n👤 ${userName}\n📅 ${date}\n📝 ${existingReason}\n\n確認取消嗎？回覆「是」確認，「否」保留。`,
    });
  }

  // ===== 查詢功能 =====
  if (/我的請假|我的紀錄/.test(text)) {
    const userName = await getUserName(event);
    let rows;
    try { rows = await getMyLeaves(userName); } catch (e) { rows = []; }
    if (rows.length === 0) {
      return client.replyMessage(event.replyToken, { type: 'text', text: '本月沒有請假紀錄。' });
    }
    const now = new Date();
    const lines = rows.map(r => `📅 ${r[1]}　📝 ${r[2] || '未說明'}`);
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: `📋 ${userName} 本月請假紀錄（${now.getMonth() + 1}月）\n\n${lines.join('\n')}`,
    });
  }

  const dateQueryMatch = text.match(/(\d{1,2})\/(\d{1,2})\s*(?:的)?(?:請假名單|有誰請假|請假|名單)/);
  if (dateQueryMatch) {
    const now = new Date();
    const date = `${now.getFullYear()}/${dateQueryMatch[1]}/${dateQueryMatch[2]}`;
    if (!isPracticeDay(date)) {
      return client.replyMessage(event.replyToken, {
        type: 'text',
        text: `${date}（週${weekdayName(date)}）沒有練習。`,
      });
    }
    let rows;
    try { rows = await getLeavesOnDate(date); } catch (e) { rows = []; }
    if (rows.length === 0) {
      return client.replyMessage(event.replyToken, { type: 'text', text: `${date} 目前沒有請假紀錄。` });
    }
    const lines = rows.map(r => `👤 ${r[0]}　📝 ${r[2] || '未說明'}`);
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: `📋 ${date} 請假名單\n\n${lines.join('\n')}`,
    });
  }

  if (/本週請假|這週請假|本週名單|這週名單/.test(text)) {
    let result;
    try { result = await getWeekLeaves(); } catch (e) { result = { dates: [], rows: [] }; }
    if (result.rows.length === 0) {
      return client.replyMessage(event.replyToken, { type: 'text', text: '本週目前沒有請假紀錄。' });
    }
    const grouped = {};
    for (const date of result.dates) grouped[date] = [];
    for (const r of result.rows) {
      if (grouped[r[1]]) grouped[r[1]].push(r[0]);
    }
    const sections = result.dates.map(date => {
      const names = grouped[date];
      if (names.length === 0) return `📅 ${date}（週${weekdayName(date)}）：無請假`;
      return `📅 ${date}（週${weekdayName(date)}）\n${names.map(n => `  👤 ${n}`).join('\n')}`;
    });
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: `📋 本週請假名單\n\n${sections.join('\n\n')}`,
    });
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

  if (date !== '日期未指定' && !isPracticeDay(date)) {
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: `${date}（週${weekdayName(date)}）沒有練習，不需要請假。\n練習日為每週二、四、六。`,
    });
  }

  const reason = extractReason(text);

  const pending = {
    type: 'leave',
    userName,
    date: date !== '日期未指定' ? date : null,
    reason: reason !== '未說明' ? reason : null,
    needsConfirm: classification === 'confirm',
  };
  pending.step = nextStep(pending);

  if (pending.step === 'ready') {
    return await checkAndRecord(event, senderId, pending);
  }

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
