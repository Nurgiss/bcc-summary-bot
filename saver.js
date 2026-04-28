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

  // GPT саммари + извлечение tensions параллельно
  let summary = '';
  let tensionsBlock = '';
  let tensionsJson = [];
  try {
    const [summaryResp, tensionsResp] = await Promise.all([
      openai.chat.completions.create({
        model: process.env.OPENAI_MODEL || 'gpt-4.1-mini',
        messages: [
          {
            role: 'system',
            content: `Ты — ассистент дизайн-команды БЦЦ. Напиши структурированное саммари встречи НА РУССКОМ, строго следуя стилю эталона.

СТРУКТУРА (только если есть данные в файле):
1. Заголовок: дата, ведущий, секретарь
2. 🏖️ Графики отпусков — кто и когда
3. ✨ Изменения в дизайн-системе
5. 📌 Пересечения и решения — подробно, это важно
6. 📢 Новости и обучение
7. 🙏 Благодарности

ПРАВИЛА:
- Задачи дизайнеров НЕ пиши — вместо этого будет ссылка на файл
- Tensions НЕ включай — они добавляются отдельно
- Пересечения, новости, благодарности — пиши подробно
- Только HTML теги <b> и <i>, никаких <br> <ul> <div>
- Не придумывай то чего нет в файле${exampleBlock}`
          },
          {
            role: 'user',
            content: `Вот содержимое файла встречи Дизайн Круга от ${dateTag}:\n\n${rawText}`
          }
        ]
      }),
      openai.chat.completions.create({
        model: process.env.OPENAI_MODEL || 'gpt-4.1-mini',
        messages: [
          {
            role: 'system',
            content: `Ты — ассистент дизайн-команды. Извлеки tensions из протокола встречи Дизайн Круга.

ВАЖНО: mammoth читает таблицы Word слипая ячейки — данные одной строки таблицы идут подряд через переносы строк.

Секция "❓ Tension" — колонки в таблице: Имя → Что вызывает напряжение/вопрос → Почему важно → Возможные шаги / Кто поможет
Секция "Tension-ы с прошлой встречи" — колонки: Имя → Что вызывало → Почему важно → Шаги → Решили вопрос → Дата финала

Правила разбивки на поля:
1. imya = только имя человека (например "Нариман", "Дида") — без лишнего текста
2. vopros = суть проблемы/вопроса — только из колонки "Что вызывает напряжение"
3. pochemu = причина важности — только из колонки "Почему важно"
4. shagi = шаги или кто поможет — только из колонки "Возможные шаги"
5. НЕ смешивай содержимое полей между собой
6. НЕ добавляй ничего от себя — только то что есть в тексте
7. Если данных нет — пустая строка ""
8. Для прошлых tensions: reshili="да" если решено, иначе "нет"

Верни ТОЛЬКО валидный JSON-массив (без markdown-блоков, без пояснений):
[{"imya":"Нариман","vopros":"В Jira меньше задач чем обсуждается на стендапах","pochemu":"Создаёт ощущение что работы нет, блокеры не видны вовремя","shagi":"Любая устная договорённость фиксируется в Jira с ответственным и результатом","reshili":"нет","data":"","status":"🔴"}]

Статус: 🔴 = новое/критично, 🟡 = в процессе/обсуждается.
Если tensions нет — верни [].`
          },
          {
            role: 'user',
            content: `Протокол встречи Дизайн Круга от ${dateTag}:\n\n${rawText}`
          }
        ]
      })
    ]);

    summary = summaryResp.choices[0].message.content.trim();

    // Парсим tensions
    try {
      const raw = tensionsResp.choices[0].message.content.trim();
      const jsonMatch = raw.match(/\[[\s\S]*\]/);
      if (jsonMatch) {
        tensionsJson = JSON.parse(jsonMatch[0]).map(t => ({
            imya:    t.imya    || t.person || t.name || t['имя'] || '',
            vopros:  t.vopros  || t.tension || t.question || t['вопрос'] || '',
            pochemu: t.pochemu || t.importance || t.why || t['почему'] || '',
            shagi:   t.shagi   || (Array.isArray(t.possible_steps) ? t.possible_steps.join('; ') : t.possible_steps) || t.steps || t['шаги'] || '',
            reshili: t.reshili || t.resolved || 'нет',
            data:    t.data    || t.deadline || '',
            status:  t.status  || '🔴',
            dateAdded: new Date().toISOString()
          }));
      }
    } catch (e) {
      console.log('⚠️  Не удалось распарсить tensions JSON:', e.message);
    }

    // Формируем блок tensions для сообщения
    if (tensionsJson.length > 0) {
        const lines = tensionsJson.map(t => {
          let s = `${t.status || '🔴'} <b>${t.imya || '?'} вызывает напряжение:</b>\n${t.vopros}`;
          if (t.pochemu) s += `\n<i>Почему важно:</i> ${t.pochemu}`;
          if (t.shagi)  s += `\n<i>Возможные шаги:</i> ${t.shagi}`;
          if (t.data)   s += `\n<i>Дедлайн:</i> ${t.data}`;
          return s;
        });
        tensionsBlock = `\n\n📌 <b>Tensions встречи (${tensionsJson.length}):</b>\n\n${lines.join('\n\n')}`;
    }

  } catch (e) {
    console.error('❌ Ошибка GPT:', e.message);
    if (NOTIFY_CHAT_ID) sendTelegramMessage(NOTIFY_CHAT_ID, `❌ Ошибка GPT при генерации саммари: ${e.message}`);
    return;
  }

  // Сохраняем tensions на VPS через SSH
  if (tensionsJson.length > 0) {
    try {
      const tensionsJsonStr = JSON.stringify(tensionsJson, null, 2).replace(/'/g, "'\\''");
      execSync(
        `sshpass -p '${VPS_PASS}' ssh -o StrictHostKeyChecking=no ${VPS_HOST} "echo '${tensionsJsonStr}' > /home/ubuntu/bcc-bot/tensions.json"`,
        { timeout: 10000 }
      );
      console.log(`📌 Сохранено ${tensionsJson.length} tensions на VPS`);
    } catch (e) {
      console.error('⚠️  Не удалось сохранить tensions на VPS:', e.message);
    }
  }

  const fullPreview = summary + tensionsBlock;

  // Сохраняем pending_summary для index.js (обработка кнопок подтверждения)
  const pendingPath = path.join(__dirname, 'pending_summary.json');
  fs.writeFileSync(pendingPath, JSON.stringify({
    summary: fullPreview,     // полный текст (саммари + tensions) — идёт в группу
    summaryOnly: summary,      // только саммари — для регенерации с дополнением
    tensionsBlock,
    tensionsJson,
    rawText: rawText,
    dateTag,
    exampleBlock
  }, null, 2));

  // Шлём только в личку на проверку — в группу отправит админ через бот
  const previewButtons = {
    inline_keyboard: [
      [{ text: '✅ Отправить в группу', callback_data: 'send_summary_to_group' }],
      [{ text: '➕ Добавить к итогам', callback_data: 'add_to_summary' }],
      [{ text: '❌ Отменить', callback_data: 'cancel_summary' }]
    ]
  };
  sendTelegramMessage(NOTIFY_CHAT_ID, `👁 <b>Превью саммари встречи ${dateTag}:</b>\n\n${fullPreview}`, previewButtons);
  console.log(`✅ Саммари сохранено, отправлено на проверку в ${NOTIFY_CHAT_ID}`);
}

// Вторник 13:30 Asia/Almaty
cron.schedule('30 13 * * 2', () => {
  generateMeetingSummary().catch(e => console.error('Саммари крон ошибка:', e.message));
}, { timezone: MEETING_TZ });

console.log('📝 Авто-саммари: каждый вторник в 13:30\n');

checkAndSync();
setInterval(checkAndSync, CHECK_INTERVAL_MS);
