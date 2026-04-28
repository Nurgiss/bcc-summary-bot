const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '.env') });
const { execSync } = require('child_process');
const fs = require('fs');
const https = require('https');
const { execFileSync } = require('child_process');
const cron = require('node-cron');
const mammoth = require('mammoth');
const OpenAI = require('openai');

const BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const VPS_HOST = 'ubuntu@82.115.43.205';
const VPS_PASS = process.env.VPS_PASSWORD || 'Lvzfwhdf5inidiynitz@';
const VPS_OUTPUT = '/home/ubuntu/bcc-bot/output/';
const ONEDRIVE_DIR = path.join(process.env.HOME, 'Library/CloudStorage/OneDrive-BCC/Meetings/Дизайн Круга');
const PROCESSED_FILE = path.join(__dirname, 'processed_files.json');
const NOTIFY_CHAT_ID = process.env.NOTIFY_CHAT_ID;
const GROUP_CHAT_ID = process.env.GROUP_CHAT_ID;
const SHAREPOINT_BASE = 'https://bcckz0-my.sharepoint.com/personal/nurgissa_anuarbek_bcc_kz';
const CHECK_INTERVAL_MS = 2 * 60 * 1000;
const MEETING_TZ = 'Asia/Almaty';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// Строим ссылку через sourcedoc GUID из БД OneDrive (надёжный способ)
const ONEDRIVE_DB = path.join(
  process.env.HOME,
  'Library/Containers/com.microsoft.OneDrive-mac/Data/Library/Application Support/OneDrive/settings/Business1/SyncEngineDatabase.db'
);

function buildFileUrl(fileName) {
  try {
    const result = execFileSync('sqlite3', [
      ONEDRIVE_DB,
      `SELECT eTag FROM od_ClientFile_Records WHERE fileName = '${fileName.replace(/'/g, "''")}';`
    ], { encoding: 'utf8', timeout: 5000 }).trim();
    // eTag формат: "{E8E99BF8-E7CE-41A8-9098-5B4F25C988AE},2"
    const match = result.match(/\{([A-F0-9-]+)\}/i);
    if (match) {
      const guid = match[1];
      console.log(`🔗 GUID для ${fileName}: ${guid}`);
      return `${SHAREPOINT_BASE}/_layouts/15/Doc.aspx?sourcedoc={${guid}}&action=default`;
    }
  } catch (e) {
    console.error('❌ Не удалось получить GUID из БД OneDrive:', e.message);
  }
  // fallback — ссылка на папку
  return `${SHAREPOINT_BASE}/_layouts/15/onedrive.aspx`;
}

// Убеждаемся что папки существуют
if (!fs.existsSync(ONEDRIVE_DIR)) fs.mkdirSync(ONEDRIVE_DIR, { recursive: true });

function loadProcessed() {
  try { return JSON.parse(fs.readFileSync(PROCESSED_FILE, 'utf8')); }
  catch { return []; }
}

function saveProcessed(list) {
  fs.writeFileSync(PROCESSED_FILE, JSON.stringify(list, null, 2));
}

function sendTelegramMessage(chatId, text, replyMarkup = null) {
  if (!BOT_TOKEN || !chatId) {
    console.error('❌ Нет BOT_TOKEN или chatId, уведомление не отправлено');
    return;
  }
  const url = `https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`;
  const payload = { chat_id: chatId, text, parse_mode: 'HTML', disable_web_page_preview: true };
  if (replyMarkup) payload.reply_markup = replyMarkup;
  const data = JSON.stringify(payload);
  const req = https.request(url, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) } }, (res) => {
    let body = '';
    res.on('data', d => body += d);
    res.on('end', () => {
      const r = JSON.parse(body);
      if (!r.ok) console.error('❌ Telegram ошибка:', r.description);
      else console.log('📨 Уведомление отправлено в Telegram');
    });
  });
  req.on('error', e => console.error('❌ Ошибка отправки в Telegram:', e.message));
  req.write(data);
  req.end();
}

async function checkAndSync() {
  try {
    // Получаем список файлов на VPS
    const result = execSync(
      `sshpass -p '${VPS_PASS}' ssh -o StrictHostKeyChecking=no ${VPS_HOST} 'ls ${VPS_OUTPUT} 2>/dev/null'`,
      { encoding: 'utf8', timeout: 15000 }
    ).trim();

    if (!result) {
      console.log(`[${new Date().toLocaleTimeString()}] Нет файлов в output/`);
      return;
    }

    const remoteFiles = result.split('\n').filter(f => f.endsWith('.docx'));
    const remoteAllFiles = result.split('\n').filter(Boolean);
    const processed = loadProcessed();
    const newFiles = remoteFiles.filter(f => !processed.includes(f));

    if (newFiles.length === 0) {
      console.log(`[${new Date().toLocaleTimeString()}] Нет новых файлов`);
      return;
    }

    for (const fileName of newFiles) {
      const destPath = path.join(ONEDRIVE_DIR, fileName);

      // Проверяем дату из имени файла: "[27.04.2026] Файл встречи..."
      const dateMatch = fileName.match(/\[(\d{2}\.\d{2}\.\d{4})\]/);
      if (dateMatch) {
        const fileDate = dateMatch[1];
        const existing = fs.readdirSync(ONEDRIVE_DIR).find(f => f.includes(fileDate));
        if (existing) {
          console.log(`⏭  Дубликат (${fileDate} уже есть): ${existing}`);
          processed.push(fileName);
          saveProcessed(processed);
          sendTelegramMessage(`⚠️ <b>Дубликат!</b> Файл за <b>${fileDate}</b> уже есть в OneDrive.\n📄 ${existing}\n\nНовый файл не сохранён.`);
          continue;
        }
      }

      console.log(`⬇️  Скачиваю: ${fileName}`);
      execSync(
        `sshpass -p '${VPS_PASS}' scp -o StrictHostKeyChecking=no '${VPS_HOST}:${VPS_OUTPUT}${fileName}' '${destPath}'`,
        { timeout: 30000 }
      );
      console.log(`✅ Сохранён в OneDrive: ${fileName}`);
      processed.push(fileName);
      saveProcessed(processed);

      // Уведомление себе — файл сохранён
      sendTelegramMessage(NOTIFY_CHAT_ID, `✅ <b>Файл сохранён в OneDrive</b>\n\n📄 ${fileName}\n📁 Meetings / Дизайн Круга`);

      // Ждём немного чтобы OneDrive зарегистрировал файл в своей БД
      try { execFileSync('sleep', ['8']); } catch(e) {}

      // Проверяем есть ли meta.json для этого файла — тогда шлём анонс
      const dateTag = fileName.match(/\[(\d{2}\.\d{2}\.\d{4})\]/)?.[1];
      if (dateTag) {
        const metaFileName = `[${dateTag}] meta.json`;
        if (remoteAllFiles.includes(metaFileName) && !processed.includes(metaFileName)) {
          try {
            const metaRaw = execSync(
              `sshpass -p '${VPS_PASS}' ssh -o StrictHostKeyChecking=no ${VPS_HOST} 'cat ${VPS_OUTPUT}${metaFileName}'`,
              { encoding: 'utf8', timeout: 10000 }
            );
            const meta = JSON.parse(metaRaw);
            const fileUrl = buildFileUrl(fileName);
            const tensionNote = meta.openTensionsCount > 0 ? `\n⚠️ Нерешённых tensions: ${meta.openTensionsCount} шт.` : '';
            const announce =
              `📢 <b>Встреча Дизайн Круга — ${meta.date}</b>\n\n` +
              `📅 Завтра в 11:30\n` +
              `👤 Ведущий: <b>${meta.vedushiy}</b>\n` +
              `📝 Секретарь: <b>${meta.sekretar}</b>` +
              tensionNote + `\n\n` +
              `📁 <a href="${fileUrl}">Открыть файл встречи</a>`;

            // Отправить в группу (или пока себе если GROUP_CHAT_ID не переключён)
            const targetChat = GROUP_CHAT_ID || NOTIFY_CHAT_ID;
            sendTelegramMessage(targetChat, announce);
            console.log(`📢 Анонс отправлен для ${dateTag}`);
            processed.push(metaFileName);
            saveProcessed(processed);
          } catch (e) {
            console.error(`❌ Ошибка чтения meta.json для ${dateTag}:`, e.message);
          }
        }
      }
    }
  } catch (e) {
    console.error(`[${new Date().toLocaleTimeString()}] Ошибка:`, e.message);
  }
}

console.log('🔄 BCC OneDrive Saver запущен');
console.log(`📁 OneDrive папка: ${ONEDRIVE_DIR}`);
console.log(`⏱  Проверка каждые ${CHECK_INTERVAL_MS / 60000} мин\n`);

// ─── Авто-саммари: вторник 13:30 ─────────────────────────────────────────────
async function generateMeetingSummary() {
  console.log('📝 Запускаю авто-саммари встречи...');

  // Найти файл встречи сегодня (вторник)
  const now = new Date(new Date().toLocaleString('en-US', { timeZone: MEETING_TZ }));
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  const dateTag = `${dd}.${mm}.${yyyy}`;
  const fileName = `[${dateTag}] Файл встречи Дизайн Круга.docx`;
  const filePath = path.join(ONEDRIVE_DIR, fileName);

  if (!fs.existsSync(filePath)) {
    console.log(`⚠️  Файл не найден для саммари: ${fileName}`);
    if (NOTIFY_CHAT_ID) sendTelegramMessage(NOTIFY_CHAT_ID, `⚠️ Авто-саммари: файл <b>${fileName}</b> не найден в OneDrive`);
    return;
  }

  // Парсим docx → текст
  let rawText = '';
  try {
    const result = await mammoth.extractRawText({ path: filePath });
    rawText = result.value.trim();
  } catch (e) {
    console.error('❌ Ошибка парсинга docx:', e.message);
    return;
  }

  if (rawText.length < 200) {
    console.log('⚠️  Файл почти пустой, саммари не делаем');
    if (NOTIFY_CHAT_ID) sendTelegramMessage(NOTIFY_CHAT_ID, `⚠️ Файл встречи <b>${dateTag}</b> почти пустой — саммари не создано.\nВозможно секретарь ещё не заполнил.`);
    return;
  }

  // Получаем эталон с VPS
  let exampleText = '';
  try {
    exampleText = execSync(
      `sshpass -p '${VPS_PASS}' ssh -o StrictHostKeyChecking=no ${VPS_HOST} 'cat /home/ubuntu/bcc-bot/example_protocol.txt 2>/dev/null || echo ""'`,
      { encoding: 'utf8', timeout: 10000 }
    ).trim();
    if (exampleText) console.log(`📚 Эталон загружен с VPS (${exampleText.length} символов)`);
    else console.log('ℹ️  Эталон на VPS не найден, используем встроенный стиль');
  } catch (e) {
    console.log('ℹ️  Не удалось загрузить эталон с VPS:', e.message);
  }

  const exampleBlock = exampleText
    ? `\n\nЭТАЛОН СТИЛЯ И СТРУКТУРЫ (строго следуй этому формату):\n${exampleText}`
    : '';

  // GPT саммари
  let summary = '';
  try {
    const completion = await openai.chat.completions.create({
      model: process.env.OPENAI_MODEL || 'gpt-4.1-mini',
      messages: [
        {
          role: 'system',
          content: `Ты — ассистент дизайн-команды. Напиши краткое и чёткое summary встречи на русском языке.\nСтрого следуй стилю, структуре и тону эталона ниже.\nИспользуй HTML теги для форматирования (<b>, <i>). Максимум 800 символов.${exampleBlock}`
        },
        {
          role: 'user',
          content: `Вот содержимое файла встречи Дизайн Круга от ${dateTag}:\n\n${rawText.substring(0, 6000)}`
        }
      ]
    });
    summary = completion.choices[0].message.content.trim();
  } catch (e) {
    console.error('❌ Ошибка GPT:', e.message);
    if (NOTIFY_CHAT_ID) sendTelegramMessage(NOTIFY_CHAT_ID, `❌ Ошибка GPT при генерации саммари: ${e.message}`);
    return;
  }

  // Сохраняем pending_summary для index.js (обработка кнопок подтверждения)
  const pendingPath = path.join(__dirname, 'pending_summary.json');
  fs.writeFileSync(pendingPath, JSON.stringify({ summary, rawText: rawText.substring(0, 6000), dateTag, exampleBlock }, null, 2));

  // Шлём только в личку на проверку — в группу отправит админ через бот
  const previewButtons = {
    inline_keyboard: [
      [{ text: '✅ Отправить в группу', callback_data: 'send_summary_to_group' }],
      [{ text: '➕ Добавить к итогам', callback_data: 'add_to_summary' }],
      [{ text: '❌ Отменить', callback_data: 'cancel_summary' }]
    ]
  };
  sendTelegramMessage(NOTIFY_CHAT_ID, `👁 <b>Превью саммари встречи ${dateTag}:</b>\n\n${summary}`, previewButtons);
  console.log(`✅ Саммари сохранено, отправлено на проверку в ${NOTIFY_CHAT_ID}`);
}

// Вторник 13:30 Asia/Almaty
cron.schedule('30 13 * * 2', () => {
  generateMeetingSummary().catch(e => console.error('Саммари крон ошибка:', e.message));
}, { timezone: MEETING_TZ });

console.log('📝 Авто-саммари: каждый вторник в 13:30\n');

checkAndSync();
setInterval(checkAndSync, CHECK_INTERVAL_MS);
