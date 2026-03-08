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

// ===== 請假關鍵字 =====
const LEAVE_KEYWORDS = [
  '請假', '缺席', '不來', '無法出席', '不能來', '沒辦法來',
  '有事', '不能參加', '無法參加', '請個假', '這次不來',
  '今天不來', '明天不來', '這週不來', '下週不來',
];

function isLeaveMessage(text) {
  return LEAVE_KEYWORDS.some(kw => text.includes(kw));
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

async function handleEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') return null;

  const text = event.message.text.trim();
  const senderId = event.source.userId;

  if (!isLeaveMessage(text)) return null;

  let userName = '未知成員';
  try {
    if (event.source.type === 'group') {
      const profile = await client.getGroupMemberProfile(event.source.groupId, senderId);
      userName = profile.displayName;
    } else {
      const profile = await client.getProfile(senderId);
      userName = profile.displayName;
    }
  } catch (e) {
    console.error('取得用戶名稱失敗:', e.message);
  }

  const date = extractDate(text);
  const reason = extractReason(text);
  const now = new Date();
  const timestamp = `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()} ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;

  try {
    await appendToSheet(userName, date, reason, timestamp);
    const replyMsg = `✅ 已登記請假\n👤 ${userName}\n📅 ${date}\n📝 ${reason === '未說明' ? '原因未說明' : reason}`;
    return client.replyMessage(event.replyToken, { type: 'text', text: replyMsg });
  } catch (e) {
    console.error('處理失敗:', e.message);
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: `⚠️ 登記失敗，請管理員手動記錄。`,
    });
  }
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
