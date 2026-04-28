require('dotenv').config();

const { Telegraf, Markup } = require('telegraf');
const OpenAI = require('openai');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const _pdfParseLib = require('pdf-parse');
const pdfParse = typeof _pdfParseLib === 'function' ? _pdfParseLib : (_pdfParseLib.default || _pdfParseLib);
const mammoth = require('mammoth');
const PizZip = require('pizzip');

const cron = require('node-cron');

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4.1-mini';
const GROUP_CHAT_ID = process.env.GROUP_CHAT_ID;
const NOTIFY_CHAT_ID = process.env.NOTIFY_CHAT_ID;
const OUTPUT_DIR = path.join(__dirname, 'output');
const MEETING_HOUR = parseInt(process.env.MEETING_HOUR || '12');
const MEETING_TZ = process.env.MEETING_TIMEZONE || 'Asia/Almaty';
const SHAREPOINT_BASE = process.env.SHAREPOINT_BASE || 'https://bcckz0-my.sharepoint.com/personal/nurgissa_anuarbek_bcc_kz';
const ONEDRIVE_ROOT_PATH = process.env.ONEDRIVE_ROOT_PATH || '/personal/nurgissa_anuarbek_bcc_kz';
const ONEDRIVE_MEETINGS_SUBPATH = process.env.ONEDRIVE_MEETINGS_SUBPATH || 'Meetings/Дизайн Круга';

if (!TELEGRAM_BOT_TOKEN) throw new Error('Не найден TELEGRAM_BOT_TOKEN в .env');
if (!OPENAI_API_KEY) throw new Error('Не найден OPENAI_API_KEY в .env');

const bot = new Telegraf(TELEGRAM_BOT_TOKEN);
const openai = new OpenAI({ apiKey: OPENAI_API_KEY });

// ─── Хранение примера протокола на диске ─────────────────────────────────────
const EXAMPLE_FILE = path.join(__dirname, 'example_protocol.txt');
const PENDING_SUMMARY_FILE = path.join(__dirname, 'pending_summary.json');

function loadPendingSummary() {
  try { return JSON.parse(fs.readFileSync(PENDING_SUMMARY_FILE, 'utf8')); }
  catch { return null; }
}
function clearPendingSummary() {
  try { fs.unlinkSync(PENDING_SUMMARY_FILE); } catch {}
}

const DEFAULT_EXAMPLE = `📎 Протокол встречи дизайнеров 15 апреля 2025

Текущие задачи и планы в таблице: [Открыть](https://notion.so/example)

КРАТКОЕ СОДЕРЖАНИЕ
Встреча прошла продуктивно: обсудили пересечения отпусков на май, разобрали накопившиеся tensions и договорились о приоритетах на следующие две недели. Отдельно отметили вклад Алины в синхронизацию команды.

КЛЮЧЕВЫЕ МОМЕНТЫ

🌞 Графики отпусков:
— Саша: 1–10 мая
— Дима: 20–25 мая
— Пересечение с релизом — договорились прикрыть Сашину зону

🛠 Пересечения и решения:
— Два дизайнера параллельно работали над онбордингом — разделили зоны ответственности
— Компонент кнопки задублировался в двух макетах — оставили версию из дизайн-системы

📌 Tensions:

Новые:
🔴 Нет единого процесса сдачи макетов разработке
🟡 Figma-файл проекта X не структурирован — сложно передавать

Прошлые:
✅ Договорились о шаблоне именования слоёв — внедрили
🟡 Ревью макетов до передачи в разработку — в процессе

🙏 Благодарности:
— Алина — за инициативу с синхронизацией графиков
— Марат — за быстрый фидбек по компонентам`;

function loadExample() {
  try {
    if (fs.existsSync(EXAMPLE_FILE)) return fs.readFileSync(EXAMPLE_FILE, 'utf8').trim();
  } catch (e) { console.error('Ошибка чтения примера:', e.message); }
  return DEFAULT_EXAMPLE;
}

function saveExample(text) {
  try { fs.writeFileSync(EXAMPLE_FILE, text, 'utf8'); }
  catch (e) { console.error('Ошибка сохранения примера:', e.message); }
}

// ─── Парсинг файлов ───────────────────────────────────────────────────────────
async function downloadFile(fileUrl) {
  const response = await axios.get(fileUrl, { responseType: 'arraybuffer' });
  return Buffer.from(response.data);
}

async function extractTextFromBuffer(buffer, mimeType, fileName) {
  const ext = (fileName || '').toLowerCase().split('.').pop();

  if (mimeType === 'application/pdf' || ext === 'pdf') {
    const data = await pdfParse(buffer);
    return data.text.trim();
  }

  if (
    mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
    ext === 'docx'
  ) {
    const result = await mammoth.extractRawText({ buffer });
    return result.value.trim();
  }

  throw new Error('Формат не поддерживается. Пришли PDF или DOCX.');
}

async function getTextFromTelegramDocument(ctx, document) {
  const file = await ctx.telegram.getFile(document.file_id);
  const fileUrl = `https://api.telegram.org/file/bot${TELEGRAM_BOT_TOKEN}/${file.file_path}`;
  const buffer = await downloadFile(fileUrl);
  return extractTextFromBuffer(buffer, document.mime_type, document.file_name);
}


// ─── Генерация шаблона протокола ─────────────────────────────────────────────
const TEMPLATE_FILE = path.join(__dirname, 'Шаблон протокола встречи.docx');
const ROTATION_FILE = path.join(__dirname, 'rotation.json');
const TENSIONS_FILE = path.join(__dirname, 'tensions.json');

function loadSavedTensions() {
  try {
    if (fs.existsSync(TENSIONS_FILE)) return JSON.parse(fs.readFileSync(TENSIONS_FILE, 'utf8'));
  } catch {}
  return [];
}

function saveTensionsFromProtocol(protocolText) {
  try {
    // Найти блок "Прошлые:" и "Новые:" внутри раздела Tensions
    const tensionBlock = protocolText.match(/📌 Tensions?[\s\S]*?(?=\n🙏|\n💬|$)/i)?.[0] || '';

    // Парсим строки с эмодзи статуса
    const lines = tensionBlock.split('\n').map(l => l.trim()).filter(l => /^[🔴🟡✅]/.test(l));

    const tensions = lines
      .filter(l => !/^✅/.test(l)) // убираем уже решённые
      .map(line => {
        // Формат: "🔴 Имя — описание" или "🔴 описание"
        const statusMatch = line.match(/^([🔴🟡✅])\s+(.+)/u);
        if (!statusMatch) return null;
        const body = statusMatch[2];
        const dashIdx = body.indexOf(' — ');
        if (dashIdx > -1) {
          return {
            imya: body.slice(0, dashIdx).trim(),
            vopros: body.slice(dashIdx + 3).trim(),
            pochemu: '',
            shagi: '',
            reshili: 'Нет',
            data: '',
            status: statusMatch[1],
            dateAdded: new Date().toISOString()
          };
        }
        return {
          imya: '',
          vopros: body.trim(),
          pochemu: '',
          shagi: '',
          reshili: 'Нет',
          data: '',
          status: statusMatch[1],
          dateAdded: new Date().toISOString()
        };
      })
      .filter(Boolean);

    if (tensions.length > 0) {
      fs.writeFileSync(TENSIONS_FILE, JSON.stringify(tensions, null, 2), 'utf8');
      console.log(`Сохранено ${tensions.length} tensions`);
    }
  } catch (e) {
    console.error('Ошибка парсинга tensions:', e.message);
  }
}

function loadRotation() {
  try { return JSON.parse(fs.readFileSync(ROTATION_FILE, 'utf8')); }
  catch { return { team: [], history: [] }; }
}

function saveRotation(data) {
  fs.writeFileSync(ROTATION_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// Возвращает {vedushiy, sekretar} на текущий месяц, не повторяя последних
function getMonthlyAssignment() {
  const data = loadRotation();
  if (data.team.length < 2) return null;

  const monthKey = new Date().toISOString().slice(0, 7);

  const existing = data.history.find(h => h.month === monthKey);
  if (existing) return existing;

  const team = data.team;
  const restrictions = data.roleRestrictions || {}; // { "Азамат": "sekretar", "Абай": "vedushiy" }
  // Исключаем тех кто был ведущим последние N раз (не более половины команды)
  const excludeCount = Math.min(2, Math.floor(team.length / 2));
  const lastVedushie = data.history.slice(-excludeCount).map(h => h.vedushiy);
  const lastSekretari = data.history.slice(-excludeCount).map(h => h.sekretar);

  // Ведущий: не был ведущим последние N раз + не имеет запрета на роль ведущего
  const vedCandidates = team.filter(p => !lastVedushie.includes(p) && restrictions[p] !== 'vedushiy');
  const vedFallback = team.filter(p => restrictions[p] !== 'vedushiy');
  const vedushiy = (vedCandidates.length > 0 ? vedCandidates : vedFallback.length > 0 ? vedFallback : team)
    .reduce((_, __, ___, arr) => arr[Math.floor(Math.random() * arr.length)]);

  // Секретарь: не ведущий, не был секретарём последние N раз + не имеет запрета на роль секретаря
  const sekCandidates = team.filter(p => p !== vedushiy && !lastSekretari.includes(p) && restrictions[p] !== 'sekretar');
  const sekFallback = team.filter(p => p !== vedushiy && restrictions[p] !== 'sekretar');
  const sekUltra = team.filter(p => p !== vedushiy);
  const sekretar = (sekCandidates.length > 0 ? sekCandidates : sekFallback.length > 0 ? sekFallback : sekUltra)
    .reduce((_, __, ___, arr) => arr[Math.floor(Math.random() * arr.length)]);

  const assignment = { month: monthKey, vedushiy, sekretar };
  data.history.push(assignment);
  if (data.history.length > 24) data.history.shift();
  saveRotation(data);
  return assignment;
}

function templateDoneMenu() {
  return Markup.inlineKeyboard([
    [Markup.button.callback('📢 Отправить анонс в группу', 'send_template_announce')],
    [Markup.button.callback('🏠 Главное меню', 'go_main_menu')]
  ]);
}

function fillProtocolTemplate({ date, vedushiy, sekretar, tensions }) {
  const zip = new PizZip(fs.readFileSync(TEMPLATE_FILE));
  let xml = zip.files['word/document.xml'].asText();

  // [дата протокола] — в одном run
  xml = xml.replace(/\[дата протокола\]/g, date);

  // [дд.мм.гг] — разбит Word на три run
  xml = xml.replace(
    /<w:t>\[<\/w:t><\/w:r><w:r><w:t>дд\.мм\.гг<\/w:t><\/w:r><w:r><w:rPr><w:lang w:val="en-US" \/><\/w:rPr><w:t>\]<\/w:t>/,
    `<w:t>${date}</w:t>`
  );

  // [ведущий] — разбит аналогично
  xml = xml.replace(
    /<w:t>\[<\/w:t><\/w:r><w:r><w:t>ведущий<\/w:t><\/w:r><w:r><w:rPr><w:lang w:val="en-US" \/><\/w:rPr><w:t>\]<\/w:t>/,
    `<w:t>${vedushiy}</w:t>`
  );

  // [секретарь] — разбит аналогично
  xml = xml.replace(
    /<w:t>\[<\/w:t><\/w:r><w:r><w:t>секретарь<\/w:t><\/w:r><w:r><w:rPr><w:lang w:val="en-US" \/><\/w:rPr><w:t>\]<\/w:t>/,
    `<w:t>${sekretar}</w:t>`
  );

  // Tensions с прошлой встречи — вставляем строки с 4 колонками
  if (tensions && tensions.length > 0) {
    const e = (s) => (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    const makeRow = (t) => {
      return (
        `<w:tr w:rsidR="0031F638" w:rsidTr="0031F638"><w:trPr><w:trHeight w:val="570"/></w:trPr>` +
        `<w:tc><w:tcPr><w:tcW w:w="839" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${e(t.imya)}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="1600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${e(t.vopros)}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="1100" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${e(t.pochemu)}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="1600" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${e(t.shagi)}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${e(t.reshili)}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${e(t.data)}</w:t></w:r></w:p></w:tc>` +
        `</w:tr>`
      );
    };

    const tensionRows = tensions.map(makeRow).join('');

    // Найти таблицу tensions с прошлой встречи и вставить строки после шапки
    xml = xml.replace(
      /(рошлой встречи[\s\S]*?<\/w:tr>)/,
      `$1${tensionRows}`
    );
  }

  zip.file('word/document.xml', xml);
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

const sessions = new Map();

const QUESTIONS = [
  `Дата встречи? (например: 21 апреля — если сегодня, просто напиши "сегодня")`,
  'Что добавить от себя? (контекст встречи, важные акценты, что подсветить)',
  'Ссылка на таблицу с задачами?'
];

function getSessionKey(ctx) { return `${ctx.chat.id}:${ctx.from.id}`; }
function resetSession(ctx) { sessions.delete(getSessionKey(ctx)); }

function getOrCreateSession(ctx) {
  const key = getSessionKey(ctx);
  if (!sessions.has(key)) {
    sessions.set(key, {
      stage: 'idle',
      rawText: '',
      answers: { meetingDate: '', additionalContext: '', tasksLink: '' },
      currentQuestionIndex: 0
    });
  }
  return sessions.get(key);
}

function normalizeSkip(value) {
  if (!value) return '';
  const v = value.trim();
  const skipValues = ['нет', 'не нужно', 'пропустить', '-', 'n/a', 'none'];
  return skipValues.includes(v.toLowerCase()) ? '' : v;
}

function tryParseNumberedAnswers(text) {
  const match1 = text.match(/(?:^|\n)\s*1[\).:-]?\s*([\s\S]*?)(?=\n\s*2[\).:-]?\s*|$)/i);
  const match2 = text.match(/(?:^|\n)\s*2[\).:-]?\s*([\s\S]*?)(?=\n\s*3[\).:-]?\s*|$)/i);
  const match3 = text.match(/(?:^|\n)\s*3[\).:-]?\s*([\s\S]*)$/i);
  if (!match1 || !match2 || !match3) return null;
  return {
    meetingDate: normalizeSkip(match1[1] || ''),
    additionalContext: normalizeSkip(match2[1] || ''),
    tasksLink: normalizeSkip(match3[1] || '')
  };
}

function getTodayDate() {
  return new Date().toLocaleDateString('ru-RU', { day: 'numeric', month: 'long', year: 'numeric' });
}

const MONTHS_RU = {
  'января':1,'февраля':2,'марта':3,'апреля':4,'мая':5,'июня':6,
  'июля':7,'августа':8,'сентября':9,'октября':10,'ноября':11,'декабря':12
};

function formatDateForFileName(dateStr) {
  // "27 апреля 2026" → "27.04.2026"
  const parts = dateStr.trim().split(/\s+/);
  if (parts.length === 3) {
    const day = parts[0].padStart(2, '0');
    const month = String(MONTHS_RU[parts[1].toLowerCase()] || parts[1]).padStart(2, '0');
    const year = parts[2];
    return `${day}.${month}.${year}`;
  }
  return dateStr.replace(/\s/g, '_');
}

// ─── Меню ─────────────────────────────────────────────────────────────────────
function mainMenu() {
  return Markup.inlineKeyboard([
    [Markup.button.callback('📝 Создать протокол', 'start_protocol')],
    [Markup.button.callback('📋 Подготовить шаблон встречи', 'start_template')],
    [Markup.button.callback('📢 Отправить анонс в группу', 'announce_from_menu')],
    [Markup.button.callback('🎲 Ротация в группу', 'show_rotation')],
    [Markup.button.callback('📌 Tensions', 'tensions_menu'), Markup.button.callback('📊 Статистика', 'show_stats')],
    [Markup.button.callback('👁 Посмотреть эталон', 'view_example'), Markup.button.callback('✏️ Обновить эталон', 'update_example')]
  ]);
}

// ─── Промпт ───────────────────────────────────────────────────────────────────
function buildPrompt({ rawText, meetingDate, additionalContext, tasksLink }) {
  const today = getTodayDate();
  const dateLabel = (meetingDate && meetingDate.toLowerCase() !== 'сегодня')
    ? meetingDate
    : today;
  const linkLine = tasksLink ? `Текущие задачи и планы в таблице: [Открыть](${tasksLink})` : '';
  const example = loadExample();

  return `Ты структурируешь протокол встречи дизайнеров. Точно следуй стилю, структуре и тону примера ниже — это эталон.

ПРИМЕР ГОТОВОГО ПРОТОКОЛА (эталон стиля):
---
${example}
---

ПРАВИЛА:
— Ничего не придумывай
— Не добавляй новых фактов
— НЕ сокращай смысл — сохраняй максимум деталей, решений, обсуждений из оригинала
— Пиши развёрнуто: каждый пункт должен передавать суть так, чтобы не нужно было возвращаться к оригиналу
— Улучши читаемость, но не за счёт потери деталей
— Убери только явный шум, дубли и технические артефакты
— Полностью ИГНОРИРУЙ задачи (tasks, планируется, to-do и т.д.)
— ИГНОРИРУЙ разделы "Обзор текущих задач дизайнеров" и любые списки задач/планов
— НЕ добавляй разделы вне структуры ниже

ВАЖНО:
— Ссылку оформляй строго как: Текущие задачи и планы в таблице: [Открыть](url)
— Комментарий пользователя и акценты НЕ встраивай в КРАТКОЕ СОДЕРЖАНИЕ — выводи их ТОЛЬКО в разделе 💬 Дополнение от автора в самом конце
— КРАТКОЕ СОДЕРЖАНИЕ формируй ТОЛЬКО из исходного текста встречи

КРИТИЧЕСКИ ВАЖНО: дата протокола — строго "${dateLabel}". Не меняй её, не придумывай другую.
— 🔴 актуальный, не начат
— 🟡 в работе
— ✅ решён
— Раздели на: Новые (из этой встречи) и Прошлые (перенесённые)

СТРУКТУРА (строго, ничего лишнего):

📎 Протокол встречи дизайнеров ${dateLabel}
${linkLine}

КРАТКОЕ СОДЕРЖАНИЕ
(2–3 предложения живым языком)

КЛЮЧЕВЫЕ МОМЕНТЫ

🌞 Графики отпусков:
— ...

🛠 Пересечения и решения:
— ...

📌 Tensions:

Новые:
🔴 / �� / ✅ ...

Прошлые:
🔴 / 🟡 / ✅ ...

🙏 Благодарности:
— ...

💬 Дополнение от автора:
(сюда выводи комментарий, акценты и пожелания пользователя — если есть; если нет — раздел не выводи)

ВХОДНЫЕ ДАННЫЕ:

[Исходный текст встречи]
${rawText || '—'}

[Комментарий / контекст / акценты]
${additionalContext || '—'}

[Ссылка на таблицу]
${tasksLink || '—'}

Сформируй протокол строго по структуре выше.`;
}

// ─── Генерация ────────────────────────────────────────────────────────────────
async function generateProtocol(session) {
  const prompt = buildPrompt({
    rawText: session.rawText,
    meetingDate: session.answers.meetingDate,
    additionalContext: session.answers.additionalContext,
    tasksLink: session.answers.tasksLink
  });
  const response = await openai.responses.create({
    model: OPENAI_MODEL,
    input: [
      { role: 'system', content: 'Ты аккуратный редактор протоколов. Следуй инструкции строго.' },
      { role: 'user', content: prompt }
    ],
    temperature: 0.2
  });
  return (response.output_text || '').trim();
}

async function sendLongMessage(ctx, text) {
  const maxLen = 3900;
  // Конвертируем [текст](url) → <a href="url">текст</a> для нативных ссылок Telegram
  const toHtml = (str) => str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, '<a href="$2">$1</a>');

  const html = toHtml(text);

  if (html.length <= maxLen) {
    await ctx.reply(html, { parse_mode: 'HTML' });
    return;
  }
  let start = 0;
  while (start < html.length) {
    await ctx.reply(html.slice(start, start + maxLen), { parse_mode: 'HTML' });
    start += maxLen;
  }
}

async function sendLongMessageToChat(chatId, text) {
  const maxLen = 3900;
  const toHtml = (str) => str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, '<a href="$2">$1</a>');
  const html = toHtml(text);
  let start = 0;
  while (start < html.length) {
    await bot.telegram.sendMessage(chatId, html.slice(start, start + maxLen), { parse_mode: 'HTML' });
    start += maxLen;
  }
}

// ─── Сохранение файла в output/ для синхронизации с OneDrive ─────────────────
function getOneDriveFileUrl(fileName) {
  return `${SHAREPOINT_BASE}/:w:/r/Documents/Meetings/${encodeURI('Дизайн Круга/')}${encodeURIComponent(fileName)}?d=w&csf=1&web=1`;
}

function saveToOutput(buffer, fileName) {
  try {
    if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    fs.writeFileSync(path.join(OUTPUT_DIR, fileName), buffer);
    console.log('Файл сохранён в output:', fileName);
  } catch (e) {
    console.error('Ошибка сохранения в output:', e);
  }
}

// ─── Статистика ──────────────────────────────────────────────────────────────
function updateStats(vedushiy, sekretar) {
  const data = loadRotation();
  if (!data.stats) data.stats = {};
  [[vedushiy, 'vedushiy'], [sekretar, 'sekretar']].forEach(([name, role]) => {
    if (!name) return;
    if (!data.stats[name]) data.stats[name] = { vedushiy: 0, sekretar: 0 };
    data.stats[name][role]++;
  });
  saveRotation(data);
}

function getTodayRussianDate() {
  const months = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
  const d = new Date(new Date().toLocaleString('en-US', { timeZone: MEETING_TZ }));
  return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

// ─── Блокировка в группах (кроме /testrotation) ──────────────────────────────
bot.command('testrotation', async (ctx) => {
  if (ctx.chat?.type !== 'private') return; // только в личке
  const chatId = ctx.chat.id;
  const data = loadRotation();
  if (data.team.length < 2) { await ctx.reply('❌ Нет команды. Сначала /setteam'); return; }
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const monthsIn = ['январе','феврале','марте','апреле','мае','июне','июле','августе','сентябре','октябре','ноябре','декабре'];
  const d = new Date(new Date().toLocaleString('en-US', { timeZone: MEETING_TZ }));
  const monthName = monthsIn[d.getMonth()];
  const team = data.team;

  const msg = await ctx.reply(
    `🎲 <b>Выбираем ведущего и секретаря на ${monthName}...</b>\n\n👤 Ведущий: <i>думаю...</i>\n📝 Секретарь: <i>думаю...</i>`,
    { parse_mode: 'HTML' }
  );
  const msgId = msg.message_id;

  // Анимация — только рандом, ничего не сохраняем
  let fakeVed, fakeSek;
  for (let i = 0; i < 7; i++) {
    await sleep(i < 4 ? 500 : 700);
    fakeVed = team[Math.floor(Math.random() * team.length)];
    fakeSek = team.filter(p => p !== fakeVed)[Math.floor(Math.random() * (team.length - 1))];
    const dots = '⏳'.repeat((i % 3) + 1);
    await bot.telegram.editMessageText(chatId, msgId, undefined,
      `🎲 <b>Выбираем ведущего и секретаря на ${monthName}...</b>\n\n👤 Ведущий: ${dots} <i>${fakeVed}</i>\n📝 Секретарь: ${dots} <i>${fakeSek}</i>`,
      { parse_mode: 'HTML' }
    ).catch(() => {});
  }

  await sleep(900);
  // Показываем текущее реальное назначение (не меняем его)
  const current = (data.history || []).slice(-1)[0];
  const vedushiy = current?.vedushiy || fakeVed;
  const sekretar = current?.sekretar || fakeSek;
  await bot.telegram.editMessageText(chatId, msgId, undefined,
    `🎉 <b>Ротация на ${monthName} определена!</b>\n\n👑 Ведущий: <b>${vedushiy}</b>\n📋 Секретарь: <b>${sekretar}</b>\n\nВстречи каждый вторник в 11:30 — удачи! 💪\n\n<i>⚠️ Это тест — данные не изменены, в группу не отправлялось</i>`,
    { parse_mode: 'HTML' }
  ).catch(() => {});
});

bot.use(async (ctx, next) => {
  const chatType = ctx.chat?.type;
  if (chatType === 'group' || chatType === 'supergroup' || chatType === 'channel') return;
  return next();
});

// ─── Команды ──────────────────────────────────────────────────────────────────
bot.start(async (ctx) => {
  resetSession(ctx);
  getOrCreateSession(ctx);
  await ctx.reply('👋 Привет! Я помогаю составлять протоколы встреч дизайнеров.\n\nЧто хочешь сделать?', mainMenu());
});

bot.command('menu', async (ctx) => {
  resetSession(ctx);
  getOrCreateSession(ctx);
  await ctx.reply('Главное меню:', mainMenu());
});

bot.command('cancel', async (ctx) => {
  resetSession(ctx);
  await ctx.reply('Сброшено.', mainMenu());
});

bot.command('team', async (ctx) => {
  const data = loadRotation();
  if (data.team.length === 0) {
    await ctx.reply('Список команды пуст.\n\nОтправь список через запятую или каждого с новой строки:\n/setteam Алия, Марат, Дана, ...');
    return;
  }
  const assignment = getMonthlyAssignment();
  const month = new Date().toLocaleString('ru-RU', { month: 'long', year: 'numeric' });
  await ctx.reply(
    `👥 Команда (${data.team.length} чел.):\n${data.team.map((p,i) => `${i+1}. ${p}`).join('\n')}\n\n` +
    `📅 Назначение на ${month}:\n👤 Ведущий: ${assignment.vedushiy}\n📝 Секретарь: ${assignment.sekretar}`
  );
});

bot.command('setteam', async (ctx) => {
  const raw = ctx.message.text.replace('/setteam', '').trim();
  if (!raw) {
    await ctx.reply('Использование: /setteam Алия, Марат, Дана, Нурлан, ...');
    return;
  }
  const team = raw.split(/[,\n]+/).map(p => p.trim()).filter(Boolean);
  const data = loadRotation();
  data.team = team;
  data.history = [];
  saveRotation(data);

  await ctx.reply(
    `✅ Список команды обновлён (${team.length} чел.):\n${team.map((p,i) => `${i+1}. ${p}`).join('\n')}\n\nХочешь выбрать ведущего и секретаря на этот месяц?`,
    Markup.inlineKeyboard([
      [Markup.button.callback('🎲 Выбрать случайно', 'assign_random')],
      [Markup.button.callback('✏️ Выбрать вручную', 'assign_manual_start')]
    ])
  );
});

// /restrict Азамат=sekretar, Абай=vedushiy  — запретить роль конкретному человеку
// /restrict Азамат=none  — снять ограничение
bot.command('restrict', async (ctx) => {
  const raw = ctx.message.text.replace('/restrict', '').trim();
  if (!raw) {
    const data = loadRotation();
    const r = data.roleRestrictions || {};
    if (Object.keys(r).length === 0) {
      await ctx.reply('Ограничений нет.\n\nФормат: /restrict Азамат=sekretar, Абай=vedushiy\nСнять: /restrict Азамат=none');
      return;
    }
    const lines = Object.entries(r).map(([n, role]) => `• ${n}: не может быть ${role === 'sekretar' ? 'секретарём' : 'ведущим'}`);
    await ctx.reply(`🚫 <b>Ограничения ролей:</b>\n${lines.join('\n')}\n\nСнять: /restrict Имя=none`, { parse_mode: 'HTML' });
    return;
  }
  const data = loadRotation();
  if (!data.roleRestrictions) data.roleRestrictions = {};
  const entries = raw.split(/[,\n]+/).map(s => s.trim()).filter(Boolean);
  const results = [];
  for (const entry of entries) {
    const [name, role] = entry.split('=').map(s => s.trim());
    if (!name || !role) continue;
    if (role === 'none') {
      delete data.roleRestrictions[name];
      results.push(`✅ ${name}: ограничение снято`);
    } else if (role === 'sekretar' || role === 'vedushiy') {
      data.roleRestrictions[name] = role;
      results.push(`🚫 ${name}: не может быть ${role === 'sekretar' ? 'секретарём' : 'ведущим'}`);
    } else {
      results.push(`❌ ${name}: неверная роль (sekretar или vedushiy)`);
    }
  }
  saveRotation(data);
  await ctx.reply(results.join('\n'));
});

bot.action('assign_random', async (ctx) => {
  await ctx.answerCbQuery();
  const assignment = getMonthlyAssignment();
  if (!assignment) { await ctx.reply('Список команды пуст.'); return; }
  await ctx.reply(`✅ Назначено на этот месяц:\n👤 Ведущий: *${assignment.vedushiy}*\n📝 Секретарь: *${assignment.sekretar}*`, { parse_mode: 'Markdown' });
});

bot.action('assign_manual_start', async (ctx) => {
  await ctx.answerCbQuery();
  const data = loadRotation();
  if (data.team.length < 2) { await ctx.reply('Мало участников.'); return; }
  const session = getOrCreateSession(ctx);
  session.stage = 'assign_vedushiy';
  const buttons = data.team.map(p => [Markup.button.callback(p, `pick_ved_${p}`)]);
  await ctx.reply('👤 Выбери ведущего:', Markup.inlineKeyboard(buttons));
});

// Динамические кнопки выбора ведущего
bot.action(/^pick_ved_(.+)$/, async (ctx) => {
  await ctx.answerCbQuery();
  const vedushiy = ctx.match[1];
  const session = getOrCreateSession(ctx);
  session._manualVedushiy = vedushiy;
  session.stage = 'assign_sekretar';
  const data = loadRotation();
  const buttons = data.team.filter(p => p !== vedushiy).map(p => [Markup.button.callback(p, `pick_sek_${p}`)]);
  await ctx.reply(`👤 Ведущий: *${vedushiy}*\n\n📝 Выбери секретаря:`, { parse_mode: 'Markdown', ...Markup.inlineKeyboard(buttons) });
});

bot.action(/^pick_sek_(.+)$/, async (ctx) => {
  await ctx.answerCbQuery();
  const sekretar = ctx.match[1];
  const session = getOrCreateSession(ctx);
  const vedushiy = session._manualVedushiy;
  delete session._manualVedushiy;
  session.stage = 'idle';

  // Сохранить назначение
  const monthKey = new Date().toISOString().slice(0, 7);
  const data = loadRotation();
  data.history = data.history.filter(h => h.month !== monthKey);
  data.history.push({ month: monthKey, vedushiy, sekretar });
  saveRotation(data);

  await ctx.reply(`✅ Назначено на этот месяц:\n👤 Ведущий: *${vedushiy}*\n📝 Секретарь: *${sekretar}*`, { parse_mode: 'Markdown' });
});

// ─── Кнопки меню ──────────────────────────────────────────────────────────────
bot.action('start_protocol', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  session.stage = 'waiting_raw_text';
  session.rawText = '';
  session.answers = { meetingDate: '', additionalContext: '', tasksLink: '' };
  session.currentQuestionIndex = 0;
  await ctx.reply('📄 Отправь сырой текст встречи или прикрепи файл (PDF / DOCX):');
});

bot.action('template_confirm_assignment', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  const a = session.templateData._autoAssignment;
  session.templateData.vedushiy = a.vedushiy;
  session.templateData.sekretar = a.sekretar;
  // tensions из прошлого протокола подставляем автоматически
  session.templateData.tensions = session.templateData._savedTensions || [];
  delete session.templateData._autoAssignment;
  delete session.templateData._savedTensions;
  session.stage = 'template_q1';
  await ctx.reply('1️⃣ Введи дату встречи (например: 27 апреля 2026):');
});

bot.action('template_manual_assignment', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  delete session.templateData._autoAssignment;
  // Оставляем _savedTensions — спросим об этом после ведущего/секретаря
  session.stage = 'template_q1';
  await ctx.reply('1️⃣ Введи дату встречи (например: 27 апреля 2026):');
});

bot.action('tension_add_more', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  session.stage = 'template_tension_imya';
  await ctx.reply('👤 Имя (чей tension?):');
});

bot.action('tension_done', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  await ctx.reply('⏳ Генерирую файл...');
  try {
    const buffer = fillProtocolTemplate({ ...session.templateData });
    const fileName = `[${formatDateForFileName(session.templateData.date)}] Файл встречи Дизайн Круга.docx`;
    saveToOutput(buffer, fileName);
    
    session.stage = 'idle';
    session.lastTemplateData = { ...session.templateData };
    await ctx.reply('✅ Шаблон готов!', templateDoneMenu());
  } catch (e) {
    console.error('Ошибка генерации шаблона:', e);
    await ctx.reply('❌ Ошибка при генерации файла.', mainMenu());
  }
});

bot.action('send_template_announce', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  const td = session.lastTemplateData;
  if (!td) { await ctx.reply('❌ Данные не найдены. Сначала создай шаблон.'); return; }
  if (!GROUP_CHAT_ID) { await ctx.reply('❌ GROUP_CHAT_ID не задан в .env'); return; }
  try {
    const fileName2 = `[${formatDateForFileName(td.date)}] Файл встречи Дизайн Круга.docx`;
    const fileUrl2 = getOneDriveFileUrl(fileName2);
    const announce =
      `📢 <b>Встреча Дизайн Круга — ${td.date}</b>\n\n` +
      `📅 В 11:30\n` +
      `👤 Ведущий: <b>${td.vedushiy}</b>\n` +
      `📝 Секретарь: <b>${td.sekretar}</b>\n\n` +
      `📁 <a href="${fileUrl2}">Открыть файл в OneDrive</a>`;
    await bot.telegram.sendMessage(GROUP_CHAT_ID, announce, { parse_mode: 'HTML', disable_web_page_preview: true });
    await ctx.reply('✅ Анонс отправлен в группу!', mainMenu());
  } catch (e) {
    console.error('Ошибка отправки анонса:', e);
    await ctx.reply('❌ Не удалось отправить. Убедись что бот добавлен в группу.', mainMenu());
  }
});

bot.action('go_main_menu', async (ctx) => {
  await ctx.answerCbQuery();
  await ctx.reply('Главное меню:', mainMenu());
});

bot.action('announce_from_menu', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  session.announceData = {};
  const td = session.lastTemplateData;
  const rotation = getMonthlyAssignment();

  // Подтягиваем vedushiy/sekretar: из шаблона → из ротации → пустые
  const vedushiy = td?.vedushiy || rotation?.vedushiy || null;
  const sekretar = td?.sekretar || rotation?.sekretar || null;
  const date = td?.date || null;

  if (vedushiy && sekretar) {
    session.announceData.vedushiy = vedushiy;
    session.announceData.sekretar = sekretar;
    if (date) {
      session.announceData.date = date;
      session.stage = 'announce_intro';
      await ctx.reply(
        `📢 <b>Подготовка анонса</b>\n\n` +
        `📅 Дата: <b>${date}</b>\n` +
        `👤 Ведущий: <b>${vedushiy}</b>\n` +
        `📝 Секретарь: <b>${sekretar}</b>\n\n` +
        `Напиши вступительную фразу (например: «Напоминаем, в четверг собираемся!»)\n\nИли пропусти:`,
        { parse_mode: 'HTML', ...Markup.inlineKeyboard([[Markup.button.callback('Пропустить', 'announce_skip_intro')]]) }
      );
    } else {
      session.stage = 'announce_date';
      await ctx.reply(
        `📢 <b>Подготовка анонса</b>\n\n` +
        `👤 Ведущий: <b>${vedushiy}</b>\n` +
        `📝 Секретарь: <b>${sekretar}</b>\n\n` +
        `📅 Введи дату встречи (например: 29 апреля 2026):`,
        { parse_mode: 'HTML' }
      );
    }
  } else {
    session.announceData.vedushiy = '';
    session.announceData.sekretar = '';
    session.stage = 'announce_date';
    await ctx.reply('📢 <b>Подготовка анонса</b>\n\n📅 Введи дату встречи (например: 29 апреля 2026):', { parse_mode: 'HTML' });
  }
});

bot.action('announce_skip_intro', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  session.announceData.intro = '';
  session.stage = 'announce_link';
  await ctx.reply('🔗 Вставь ссылку на файл (OneDrive/SharePoint) или нажми «Без ссылки»:',
    Markup.inlineKeyboard([[Markup.button.callback('Без ссылки', 'announce_no_link')]])
  );
});

bot.action('announce_no_link', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  session.announceData.link = null;
  session.stage = 'idle';
  await showAnnouncePreview(ctx, session);
});

async function showAnnouncePreview(ctx, session) {
  const d = session.announceData;
  const linkLine = d.link ? `\n📄 Файл: <a href="${d.link}">Открыть</a>` : '';
  const introLine = d.intro ? `\n\n${d.intro}` : '';
  const dateLine = d.date ? ` — ${d.date}` : '';
  const preview =
    `📢 <b>Встреча Дизайн Круга${dateLine}</b>${introLine}${linkLine}\n\n` +
    `👤 Ведущий: ${d.vedushiy}\n` +
    `📝 Секретарь: ${d.sekretar}`;
  await ctx.reply(`👀 <b>Предпросмотр анонса:</b>\n\n${preview}\n\nОтправить в группу?`, {
    parse_mode: 'HTML',
    ...Markup.inlineKeyboard([
      [Markup.button.callback('✅ Отправить', 'announce_confirm_send')],
      [Markup.button.callback('❌ Отмена', 'go_main_menu')]
    ])
  });
}

bot.action('announce_confirm_send', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  const d = session.announceData;
  if (!d || !d.vedushiy) { await ctx.reply('❌ Данные не найдены.', mainMenu()); return; }
  session.announceData = {};
  try {
    const linkLine = d.link ? `\n📄 Файл: <a href="${d.link}">Открыть</a>` : '';
    const introLine = d.intro ? `\n\n${d.intro}` : '';
    const dateLine = d.date ? ` — ${d.date}` : '';
    const announce =
      `📢 <b>Встреча Дизайн Круга${dateLine}</b>${introLine}${linkLine}\n\n` +
      `👤 Ведущий: ${d.vedushiy}\n` +
      `📝 Секретарь: ${d.sekretar}`;
    await bot.telegram.sendMessage(GROUP_CHAT_ID, announce, { parse_mode: 'HTML' });
    await ctx.reply('✅ Анонс отправлен в группу!', mainMenu());
  } catch (e) {
    console.error('Ошибка анонса:', e);
    await ctx.reply('❌ Не удалось отправить.', mainMenu());
  }
});

bot.action('view_example', async (ctx) => {
  await ctx.answerCbQuery();
  const example = loadExample();
  await sendLongMessage(ctx, `Текущий эталон стиля:\n\n${example}`);
  await ctx.reply('Вернуться в меню:', mainMenu());
});

bot.action('start_template', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  session.stage = 'template_q1';
  session.templateData = {};

  // Попробуем автоназначить ведущего и секретаря
  const assignment = getMonthlyAssignment();
  if (assignment) {
    session.templateData._autoAssignment = assignment;
  }

  // Загрузить tensions с прошлой встречи
  const savedTensions = loadSavedTensions();
  session.templateData._savedTensions = savedTensions;

  let msg = '📋 Подготовка шаблона встречи\n\n1️⃣ Введи дату встречи (например: 27 апреля 2026):';

  if (assignment) {
    msg = `📋 Подготовка шаблона встречи\n\n` +
      `🗓 Назначение на этот месяц:\n` +
      `👤 Ведущий: *${assignment.vedushiy}*\n` +
      `📝 Секретарь: *${assignment.sekretar}*\n\n` +
      (savedTensions.length > 0
        ? `⚡️ Найдено *${savedTensions.length}* нерешённых tension(а) с прошлой встречи — подставлю автоматически.\n\n`
        : '') +
      `Подтвердить назначение или ввести вручную?`;

    await ctx.reply(msg, {
      parse_mode: 'Markdown',
      ...Markup.inlineKeyboard([
        [Markup.button.callback('✅ Подтвердить и продолжить', 'template_confirm_assignment')],
        [Markup.button.callback('✏️ Ввести ведущего/секретаря вручную', 'template_manual_assignment')]
      ])
    });
  } else {
    await ctx.reply(msg);
  }
});

bot.action('update_example', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  session.stage = 'waiting_example';
  await ctx.reply('✏️ Отправь новый пример готового протокола.\n\nОн станет эталоном стиля для всех будущих генераций.\n\n/cancel — отмена');
});

// ─── Обработка документов (PDF / DOCX) ───────────────────────────────────────
bot.on('document', async (ctx) => {
  const session = getOrCreateSession(ctx);
  const doc = ctx.message.document;
  const ext = (doc.file_name || '').toLowerCase().split('.').pop();
  const supported = ['pdf', 'docx'];

  if (!supported.includes(ext) &&
      doc.mime_type !== 'application/pdf' &&
      doc.mime_type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
    await ctx.reply('⚠️ Поддерживаются только PDF и DOCX файлы.');
    return;
  }

  // ── Если это заполненный файл встречи — сохраняем в output для saver.js ──
  const isMeetingFile = /Файл встречи Дизайн Круга/i.test(doc.file_name || '');
  if (isMeetingFile && ext === 'docx') {
    try {
      await ctx.reply(`📥 Получил заполненный файл встречи <b>${doc.file_name}</b>, сохраняю...`, { parse_mode: 'HTML' });
      const fileLink = await ctx.telegram.getFileLink(doc.file_id);
      const resp = await fetch(fileLink.href);
      const buffer = Buffer.from(await resp.arrayBuffer());
      saveToOutput(buffer, doc.file_name);
      const url = getOneDriveFileUrl(doc.file_name);
      await ctx.reply(
        `✅ Файл сохранён! saver.js подхватит его в течение 2 минут и синхронизирует с OneDrive.\n\n` +
        `📎 <a href="${url}">${doc.file_name}</a>`,
        { parse_mode: 'HTML', disable_web_page_preview: true, ...mainMenu() }
      );
    } catch (err) {
      console.error('Ошибка сохранения файла встречи:', err.message);
      await ctx.reply(`❌ Ошибка сохранения: ${err.message}`, mainMenu());
    }
    return;
  }


  if (session.stage !== 'waiting_raw_text' && session.stage !== 'idle') {
    await ctx.reply('Сначала нажми "📝 Создать протокол", затем пришли файл.');
    return;
  }

  try {
    await ctx.reply(`📂 Получил файл <b>${doc.file_name}</b>, извлекаю текст...`, { parse_mode: 'HTML' });
    const text = await getTextFromTelegramDocument(ctx, doc);

    if (!text || text.length < 50) {
      await ctx.reply('❌ Не удалось извлечь текст из файла. Попробуй скопировать текст вручную.');
      return;
    }

    session.rawText = text;
    session.stage = 'collecting_answers';
    session.currentQuestionIndex = 0;
    session.answers = { meetingDate: '', additionalContext: '', tasksLink: '' };

    await ctx.reply(
      `✅ Текст извлечён (${text.length} символов)\n\n` +
      'Теперь уточняющие вопросы:\n\n' +
      `Вопрос 1/3:\n${QUESTIONS[0]}`
    );
  } catch (err) {
    console.error('Ошибка парсинга файла:', err.message);
    await ctx.reply(`❌ Ошибка: ${err.message}`, mainMenu());
  }
});

// ─── Обработка текста ─────────────────────────────────────────────────────────
bot.on('text', async (ctx) => {
  const session = getOrCreateSession(ctx);
  const text = (ctx.message.text || '').trim();
  if (text.startsWith('/')) return;

  // ─── Дополнение к итогам встречи ─────────────────────────────────────────
  if (session.stage === 'waiting_summary_addition') {
    const pending = loadPendingSummary();
    if (!pending) { session.stage = 'idle'; await ctx.reply('❌ Саммари не найдено.', mainMenu()); return; }
    session.stage = 'idle';
    await ctx.reply('⏳ Перегенерирую саммари с дополнением...');
    let newSummary = '';
    try {
      const exampleText = loadExample();
      const exampleBlock = exampleText ? `\n\nЭТАЛОН СТИЛЯ:\n${exampleText}` : '';
      const completion = await openai.chat.completions.create({
        model: OPENAI_MODEL,
        messages: [
          {
            role: 'system',
            content: `Ты — ассистент дизайн-команды БЦЦ. Напиши полное структурированное саммари встречи НА РУССКОМ, строго следуя стилю и структуре эталона.

СТРУКТУРА САММАРИ (включи все секции из файла):
1. Заголовок с датой, ведущим, секретарём
2. 🏖️ Графики отпусков (если есть)
3. ✨ Изменения в дизайн-системе (если есть)
4. 👉 Обзор задач — по каждому дизайнеру отдельно
5. 📌 Пересечения и решения
6. 📢 Новости и обучение (если есть)
7. 🙏 Благодарности (если есть)

PRAVILA:
- По каждому дизайнеру отдельный пункт с именем жирным
- Tensions НЕ включай — они добавляются отдельно автоматически
- Используй только HTML теги <b>, <i>
- Не придумывай то чего нет в файле
- Учти дополнение от организатора${exampleBlock}`
          },
          {
            role: 'user',
            content: `Исходный файл встречи (${pending.dateTag}):\n\n${pending.rawText}\n\nДОПОЛНЕНИЕ ОТ ОРГАНИЗАТОРА:\n${text}`
          }
        ]
      });
      newSummary = completion.choices[0].message.content.trim();
    } catch (e) {
      await ctx.reply(`❌ Ошибка GPT: ${e.message}`, mainMenu());
      return;
    }
    // Собираем полный текст: новое саммари + сохранённые tensions
    const tensionsBlock = pending.tensionsBlock || '';
    const fullSummary = newSummary + tensionsBlock;
    // Обновляем pending с новым summary
    fs.writeFileSync(PENDING_SUMMARY_FILE, JSON.stringify({ ...pending, summary: fullSummary, summaryOnly: newSummary }, null, 2));
    const previewButtons = Markup.inlineKeyboard([
      [Markup.button.callback('✅ Отправить в группу', 'send_summary_to_group')],
      [Markup.button.callback('➕ Добавить ещё', 'add_to_summary')],
      [Markup.button.callback('❌ Отменить', 'cancel_summary')]
    ]);
    await ctx.reply(`👁 <b>Обновлённое саммари:</b>\n\n${fullSummary}`, { parse_mode: 'HTML', ...previewButtons });
    return;
  }

  try {
    // ─── Флоу шаблона встречи ─────────────────────────────────────────────
    if (session.stage === 'template_q1') {
      session.templateData.date = text;
      if (session.templateData.vedushiy) {
        // Ведущий/секретарь уже заданы автоматически
        if (session.templateData.tensions && session.templateData.tensions.length > 0) {
          // Tensions тоже есть — генерируем сразу
          await ctx.reply('⏳ Генерирую файл...');
          const buffer = fillProtocolTemplate({ ...session.templateData });
          const fileName = `[${formatDateForFileName(session.templateData.date)}] Файл встречи Дизайн Круга.docx`;
          saveToOutput(buffer, fileName);
          
          session.stage = 'idle';
          session.lastTemplateData = { ...session.templateData };
          await ctx.reply('✅ Шаблон готов!', templateDoneMenu());
        } else {
          // Tensions нет — спросим
          session.templateData.tensions = [];
          session.stage = 'template_tension_imya';
          await ctx.reply('4️⃣ Tension-ы с прошлой встречи\n\nЕсли нет — напиши "нет".\n\n👤 Имя (чей tension?):');
        }
      } else {
        session.stage = 'template_q2';
        await ctx.reply('2️⃣ Кто ведущий встречи?');
      }
      return;
    }
    if (session.stage === 'template_q2') {
      session.templateData.vedushiy = text;
      session.stage = 'template_q3';
      await ctx.reply('3️⃣ Кто секретарь встречи?');
      return;
    }
    if (session.stage === 'template_q3') {
      session.templateData.sekretar = text;
      const saved = session.templateData._savedTensions || [];
      delete session.templateData._savedTensions;
      if (saved.length > 0) {
        session.templateData.tensions = saved;
        await ctx.reply('⏳ Генерирую файл...');
        const buffer = fillProtocolTemplate({ ...session.templateData });
        const fileName = `[${formatDateForFileName(session.templateData.date)}] Встреча дизайн круга.docx`;
        
        session.stage = 'idle';
        session.lastTemplateData = { ...session.templateData };
          await ctx.reply('✅ Шаблон готов!', templateDoneMenu());
      } else {
        session.templateData.tensions = [];
        session.stage = 'template_tension_imya';
        await ctx.reply(
          '4️⃣ Tension-ы с прошлой встречи\n\nЕсли нет ни одного — напиши "нет".\n\nИначе начнём добавлять по одному.\n\n👤 Имя (чей tension?)'
        );
      }
      return;
    }
    if (session.stage === 'template_tension_imya') {
      if (text.toLowerCase() === 'нет') {
        // Нет tensions — генерируем сразу
        await ctx.reply('⏳ Генерирую файл...');
        const buffer = fillProtocolTemplate({ ...session.templateData });
        const fileName = `[${formatDateForFileName(session.templateData.date)}] Файл встречи Дизайн Круга.docx`;
        saveToOutput(buffer, fileName);
        
        session.stage = 'idle';
        session.lastTemplateData = { ...session.templateData };
        await ctx.reply('✅ Шаблон готов!', templateDoneMenu());
        return;
      }
      session.templateData._currentTension = { imya: text };
      session.stage = 'template_tension_vopros';
      await ctx.reply('💬 Что вызывало напряжение / вопрос?');
      return;
    }
    if (session.stage === 'template_tension_vopros') {
      session.templateData._currentTension.vopros = text;
      session.stage = 'template_tension_pochemu';
      await ctx.reply('🤔 Почему важно?');
      return;
    }
    if (session.stage === 'template_tension_pochemu') {
      session.templateData._currentTension.pochemu = text;
      session.stage = 'template_tension_shagi';
      await ctx.reply('🚶 Возможные шаги / Кто поможет?');
      return;
    }
    if (session.stage === 'template_tension_shagi') {
      session.templateData._currentTension.shagi = text;
      session.stage = 'template_tension_reshili';
      await ctx.reply('✅ Решили вопрос? (да / нет / частично / в процессе)');
      return;
    }
    if (session.stage === 'template_tension_reshili') {
      session.templateData._currentTension.reshili = text;
      session.stage = 'template_tension_data';
      await ctx.reply('📅 Дата финала если не решён (или "-" если решён):');
      return;
    }
    if (session.stage === 'template_tension_data') {
      session.templateData._currentTension.data = text === '-' ? '' : text;
      if (!session.templateData._currentTension.dateAdded) session.templateData._currentTension.dateAdded = new Date().toISOString();
      session.templateData.tensions.push(session.templateData._currentTension);
      delete session.templateData._currentTension;
      session.stage = 'template_tension_more';
      await ctx.reply(
        `✅ Tension добавлен (${session.templateData.tensions.length} шт.)\n\nДобавить ещё один?`,
        Markup.inlineKeyboard([
          [Markup.button.callback('➕ Добавить ещё', 'tension_add_more')],
          [Markup.button.callback('✅ Готово, генерировать файл', 'tension_done')]
        ])
      );
      return;
    }
    if (session.stage === 'template_tension_more') { return; } // ждём кнопку

    if (session.stage === 'waiting_example') {
      saveExample(text);
      session.stage = 'idle';
      await ctx.reply('✅ Эталон обновлён! Бот будет генерировать протоколы в этом стиле.', mainMenu());
      return;
    }

    if (session.stage === 'waiting_raw_text') {
      session.rawText = text;
      session.stage = 'collecting_answers';
      session.currentQuestionIndex = 0;
      await ctx.reply(
        'Принято ✅\n\nМожно ответить по шагам или одним сообщением:\n1: ...\n2: ...\n3: ...\n\n' +
        `Вопрос 1/3:\n${QUESTIONS[0]}`
      );
      return;
    }

    if (session.stage === 'collecting_answers') {
      if (session.currentQuestionIndex === 0) {
        const parsed = tryParseNumberedAnswers(text);
        if (parsed) { session.answers = parsed; session.currentQuestionIndex = 3; }
      }

      if (session.currentQuestionIndex < 3) {
        if (session.currentQuestionIndex === 0) {
          session.answers.meetingDate = normalizeSkip(text);
          session.currentQuestionIndex = 1;
          await ctx.reply(`Вопрос 2/3:\n${QUESTIONS[1]}`);
          return;
        }
        if (session.currentQuestionIndex === 1) {
          session.answers.additionalContext = normalizeSkip(text);
          session.currentQuestionIndex = 2;
          await ctx.reply(`Вопрос 3/3:\n${QUESTIONS[2]}`);
          return;
        }
        if (session.currentQuestionIndex === 2) {
          session.answers.tasksLink = normalizeSkip(text);
          session.currentQuestionIndex = 3;
        }
      }

      if (session.currentQuestionIndex >= 3) {
        session.stage = 'generating';
        await ctx.reply('⏳ Генерирую протокол...');
        const finalProtocol = await generateProtocol(session);
        if (!finalProtocol) throw new Error('OpenAI вернул пустой ответ');
        await sendLongMessage(ctx, finalProtocol);
        saveTensionsFromProtocol(finalProtocol);
        session.lastProtocol = finalProtocol;
        session.stage = 'idle';
        await ctx.reply('✅ Готово! Что дальше?', Markup.inlineKeyboard([
          [Markup.button.callback('📤 Отправить в группу', 'send_to_group')],
          [Markup.button.callback('📝 Новый протокол', 'start_protocol')],
        ]));
        return;
      }
    }

    if (session.stage === 'announce_vedushiy') {
      session.announceData.vedushiy = ctx.message.text.trim();
      session.stage = 'announce_sekretar';
      await ctx.reply('📝 Введи имя секретаря:');
      return;
    }

    if (session.stage === 'announce_sekretar') {
      session.announceData.sekretar = ctx.message.text.trim();
      session.stage = 'announce_date';
      await ctx.reply('📅 Введи дату встречи (например: 29 апреля 2026):');
      return;
    }

    if (session.stage === 'announce_date') {
      session.announceData.date = ctx.message.text.trim();
      session.stage = 'announce_intro';
      await ctx.reply(
        '✍️ Напиши вступительную фразу для анонса\n(например: «Напоминаем, в четверг собираемся!»)\n\nИли пропусти:',
        Markup.inlineKeyboard([[Markup.button.callback('Пропустить', 'announce_skip_intro')]])
      );
      return;
    }

    if (session.stage === 'announce_intro') {
      session.announceData.intro = ctx.message.text.trim();
      session.stage = 'announce_link';
      await ctx.reply('🔗 Вставь ссылку на файл (OneDrive/SharePoint) или нажми «Без ссылки»:',
        Markup.inlineKeyboard([[Markup.button.callback('Без ссылки', 'announce_no_link')]])
      );
      return;
    }

    if (session.stage === 'announce_link') {
      session.announceData.link = ctx.message.text.trim();
      session.stage = 'idle';
      await showAnnouncePreview(ctx, session);
      return;
    }

    if (session.stage === 'waiting_announce_link') {
      const link = ctx.message.text.trim();
      session.stage = 'idle';
      const td = session.lastTemplateData;
      if (!td) { await ctx.reply('❌ Данные не найдены.', mainMenu()); return; }
      const fileName = `[${formatDateForFileName(td.date)}] Файл встречи Дизайн Круга.docx`;
      session.pendingAnnounce = { td, link };
      const preview =
        `📋 <b>Файл встречи Дизайн Круга — ${td.date}</b>\n\n` +
        `📄 Файл: <a href="${link}">${fileName}</a>\n\n` +
        `👤 Ведущий: ${td.vedushiy}\n` +
        `📝 Секретарь: ${td.sekretar}`;
      await ctx.reply(`👀 <b>Предпросмотр анонса:</b>\n\n${preview}\n\nОтправить в группу?`, {
        parse_mode: 'HTML',
        ...Markup.inlineKeyboard([
          [Markup.button.callback('✅ Отправить в группу', 'announce_confirm_send')],
          [Markup.button.callback('❌ Отмена', 'go_main_menu')]
        ])
      });
      return;
    }

    if (session.stage === 'generating') { await ctx.reply('Секунду, ещё генерирую...'); return; }
    if (session.stage === 'idle') { await ctx.reply('Выбери действие:', mainMenu()); }

  } catch (error) {
    console.error('Ошибка:', error);
    session.stage = 'idle';
    await ctx.reply('❌ Произошла ошибка. Попробуй снова.', mainMenu());
  }
});

bot.action('send_to_group', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  if (!session.lastProtocol) {
    await ctx.reply('❌ Протокол не найден. Сначала сгенерируй протокол.');
    return;
  }
  if (!GROUP_CHAT_ID) {
    await ctx.reply('❌ GROUP_CHAT_ID не задан в .env');
    return;
  }
  try {
    await sendLongMessageToChat(GROUP_CHAT_ID, session.lastProtocol);
    await ctx.reply('✅ Протокол отправлен в группу!', mainMenu());
  } catch (e) {
    console.error('Ошибка отправки в группу:', e);
    await ctx.reply('❌ Не удалось отправить. Убедись, что бот добавлен в группу.');
  }
});

// ─── Закрытие tensions ───────────────────────────────────────────────────────
// ─── Tensions меню ───────────────────────────────────────────────────────────
function tensionsListMessage(tensions) {
  if (tensions.length === 0) return '✅ Открытых tensions нет!';
  return tensions.map((t, i) => {
    const status = t.status || '🔴';
    const imya = t.imya ? `<b>${t.imya}</b>: ` : '';
    return `${status} ${imya}${t.vopros}`;
  }).join('\n\n');
}

function tensionsListButtons(tensions) {
  const buttons = tensions.map((t, i) => {
    const status = t.status || '🔴';
    const imya = t.imya ? `${t.imya}: ` : '';
    const label = `${status} ${imya}${(t.vopros || '').slice(0, 40)}`;
    return [Markup.button.callback(label, `tension_detail_${i}`)];
  });
  buttons.push([Markup.button.callback('🏠 Главное меню', 'go_main_menu')]);
  return Markup.inlineKeyboard(buttons);
}

bot.action('tensions_menu', async (ctx) => {
  await ctx.answerCbQuery();
  const tensions = loadSavedTensions().filter(t => t.reshili !== 'да' && t.reshili !== 'Да');
  if (tensions.length === 0) {
    await ctx.reply('✅ Открытых tensions нет!', mainMenu());
    return;
  }
  const text = `📌 <b>Tensions (${tensions.length} открытых):</b>\n\n${tensionsListMessage(tensions)}\n\n👆 Нажми на tension чтобы открыть карточку`;
  await ctx.reply(text, { parse_mode: 'HTML', ...tensionsListButtons(tensions) });
});

// Обратная совместимость
bot.action('close_tensions_menu', async (ctx) => {
  await ctx.answerCbQuery();
  ctx.callbackQuery.data = 'tensions_menu';
  const tensions = loadSavedTensions().filter(t => t.reshili !== 'да' && t.reshili !== 'Да');
  if (tensions.length === 0) { await ctx.reply('✅ Открытых tensions нет!', mainMenu()); return; }
  const text = `📌 <b>Tensions (${tensions.length} открытых):</b>\n\n${tensionsListMessage(tensions)}\n\n👆 Нажми на tension чтобы открыть карточку`;
  await ctx.reply(text, { parse_mode: 'HTML', ...tensionsListButtons(tensions) });
});

bot.action(/^tension_detail_(\d+)$/, async (ctx) => {
  await ctx.answerCbQuery();
  const idx = parseInt(ctx.match[1]);
  const tensions = loadSavedTensions().filter(t => t.reshili !== 'да' && t.reshili !== 'Да');
  const t = tensions[idx];
  if (!t) { await ctx.reply('❌ Tension не найден.', mainMenu()); return; }
  const dateAdded = t.dateAdded ? new Date(t.dateAdded).toLocaleDateString('ru-RU') : '?';
  const days = t.dateAdded ? Math.floor((Date.now() - new Date(t.dateAdded)) / 86400000) : '?';
  let card = `${t.status || '🔴'} <b>${t.imya || 'Без имени'}</b>\n\n`;
  card += `📋 <b>Вопрос:</b>\n${t.vopros || '—'}\n\n`;
  if (t.pochemu) card += `🤔 <b>Почему важно:</b>\n${t.pochemu}\n\n`;
  if (t.shagi) card += `🚶 <b>Шаги / кто поможет:</b>\n${t.shagi}\n\n`;
  card += `📅 Добавлен: ${dateAdded} (${days} дн. назад)`;
  if (t.data) card += `\n⏰ Дедлайн: ${t.data}`;
  const buttons = Markup.inlineKeyboard([
    [Markup.button.callback('✅ Решено — закрыть', `t_close_${idx}`)],
    [Markup.button.callback('🟡 В процессе', `t_inprogress_${idx}`)],
    [Markup.button.callback('◀️ Назад к списку', 'tensions_menu')]
  ]);
  await ctx.reply(card, { parse_mode: 'HTML', ...buttons });
});

bot.action(/^t_close_(\d+)$/, async (ctx) => {
  await ctx.answerCbQuery();
  const idx = parseInt(ctx.match[1]);
  const all = loadSavedTensions();
  const open = all.filter(t => t.reshili !== 'да' && t.reshili !== 'Да');
  const target = open[idx];
  if (!target) { await ctx.reply('❌ Tension не найден.', mainMenu()); return; }
  const allIdx = all.findIndex(t => t.imya === target.imya && t.vopros === target.vopros);
  if (allIdx !== -1) {
    all[allIdx].reshili = 'да';
    all[allIdx].dateClosed = new Date().toISOString();
    fs.writeFileSync(TENSIONS_FILE, JSON.stringify(all, null, 2), 'utf8');
  }
  const remaining = all.filter(t => t.reshili !== 'да' && t.reshili !== 'Да');
  if (remaining.length === 0) {
    await ctx.reply(`✅ <b>${target.imya || target.vopros.slice(0,40)}</b> закрыт!\n\nВсе tensions закрыты 🎉`, { parse_mode: 'HTML', ...mainMenu() });
    return;
  }
  const text = `✅ Закрыт!\n\n📌 <b>Осталось tensions (${remaining.length}):</b>\n\n${tensionsListMessage(remaining)}`;
  await ctx.reply(text, { parse_mode: 'HTML', ...tensionsListButtons(remaining) });
});

bot.action(/^t_inprogress_(\d+)$/, async (ctx) => {
  await ctx.answerCbQuery();
  const idx = parseInt(ctx.match[1]);
  const all = loadSavedTensions();
  const open = all.filter(t => t.reshili !== 'да' && t.reshili !== 'Да');
  const target = open[idx];
  if (!target) { await ctx.reply('❌ Tension не найден.', mainMenu()); return; }
  const allIdx = all.findIndex(t => t.imya === target.imya && t.vopros === target.vopros);
  if (allIdx !== -1) {
    all[allIdx].reshili = 'в процессе';
    all[allIdx].status = '🟡';
    fs.writeFileSync(TENSIONS_FILE, JSON.stringify(all, null, 2), 'utf8');
  }
  await ctx.reply(`🟡 Отмечено «в процессе».`, { parse_mode: 'HTML', ...tensionsListButtons(all.filter(t => t.reshili !== 'да' && t.reshili !== 'Да')) });
});

bot.action(/^close_t_(\d+)$/, async (ctx) => {
  // legacy — редиректим на новый обработчик
  await ctx.answerCbQuery();
  const idx = parseInt(ctx.match[1]);
  const all = loadSavedTensions();
  const open = all.filter(t => t.reshili !== 'да' && t.reshili !== 'Да');
  const target = open[idx];
  if (!target) { await ctx.reply('❌ Tension не найден.', mainMenu()); return; }
  const allIdx = all.findIndex(t => t.imya === target.imya && t.vopros === target.vopros);
  if (allIdx !== -1) {
    all[allIdx].reshili = 'да';
    all[allIdx].dateClosed = new Date().toISOString();
    fs.writeFileSync(TENSIONS_FILE, JSON.stringify(all, null, 2), 'utf8');
  }
  const remaining = all.filter(t => t.reshili !== 'да' && t.reshili !== 'Да');
  if (remaining.length === 0) {
    await ctx.reply(`✅ Все tensions закрыты 🎉`, mainMenu());
    return;
  }
  await ctx.reply(`✅ Закрыт! Осталось: ${remaining.length}`, { parse_mode: 'HTML', ...tensionsListButtons(remaining) });
});

// ─── Добавление к саммари встречи ───────────────────────────────────────────
bot.action('send_summary_to_group', async (ctx) => {
  await ctx.answerCbQuery();
  const pending = loadPendingSummary();
  if (!pending) { await ctx.reply('❌ Саммари не найдено. Возможно уже было отправлено.', mainMenu()); return; }
  const targetChat = GROUP_CHAT_ID || NOTIFY_CHAT_ID;
  try {
    await bot.telegram.sendMessage(targetChat, pending.summary, { parse_mode: 'HTML' });
    clearPendingSummary();
    await ctx.editMessageReplyMarkup({ inline_keyboard: [] }).catch(() => {});
    await ctx.reply('✅ Саммари отправлено в группу!', mainMenu());
  } catch (e) {
    await ctx.reply(`❌ Ошибка: ${e.message}`, mainMenu());
  }
});

bot.action('add_to_summary', async (ctx) => {
  await ctx.answerCbQuery();
  const session = getOrCreateSession(ctx);
  const pending = loadPendingSummary();
  if (!pending) { await ctx.reply('❌ Саммари не найдено.', mainMenu()); return; }
  session.stage = 'waiting_summary_addition';
  await ctx.reply('💬 Напиши что добавить к итогам встречи:\n\n(Например: дополнительный пункт по задачам, важное решение или комментарий)');
});

bot.action('cancel_summary', async (ctx) => {
  await ctx.answerCbQuery();
  clearPendingSummary();
  await ctx.editMessageReplyMarkup({ inline_keyboard: [] }).catch(() => {});
  await ctx.reply('❌ Саммари отменено.', mainMenu());
});

// ─── Статистика ───────────────────────────────────────────────────────────────
bot.action('show_stats', async (ctx) => {
  await ctx.answerCbQuery();
  const data = loadRotation();
  const stats = data.stats || {};
  if (Object.keys(stats).length === 0) {
    await ctx.reply('📊 Статистика пока пуста — данные накапливаются при генерации шаблонов.', mainMenu());
    return;
  }
  const sorted = Object.entries(stats).sort((a, b) => (b[1].vedushiy + b[1].sekretar) - (a[1].vedushiy + a[1].sekretar));
  const lines = sorted.map(([name, s]) => `👤 <b>${name}</b>: ведущий ${s.vedushiy}x, секретарь ${s.sekretar}x`);
  await ctx.reply(`📊 <b>Статистика участия:</b>\n\n${lines.join('\n')}`, { parse_mode: 'HTML', ...mainMenu() });
});

// Запустить анимированную ротацию вручную
bot.command('newmonth', async (ctx) => {
  await ctx.reply('🎲 Запускаю ротацию в группе...');
  await animateRotationInGroup().catch(e => ctx.reply('❌ Ошибка: ' + e.message));
});

bot.action('show_rotation', async (ctx) => {
  await ctx.answerCbQuery();
  await ctx.reply('🎲 Запускаю ротацию в группе...');
  await animateRotationInGroup().catch(e => ctx.reply('❌ Ошибка: ' + e.message));
});

bot.command('stats', async (ctx) => {
  const data = loadRotation();
  const stats = data.stats || {};
  if (Object.keys(stats).length === 0) { await ctx.reply('📊 Статистика пока пуста.'); return; }
  const sorted = Object.entries(stats).sort((a, b) => (b[1].vedushiy + b[1].sekretar) - (a[1].vedushiy + a[1].sekretar));
  const lines = sorted.map(([name, s]) => `👤 <b>${name}</b>: ведущий ${s.vedushiy}x, секретарь ${s.sekretar}x`);
  await ctx.reply(`📊 <b>Статистика участия:</b>\n\n${lines.join('\n')}`, { parse_mode: 'HTML' });
});

// ─── Анимированная ротация ───────────────────────────────────────────────────
async function animateRotationInGroup() {
  if (!GROUP_CHAT_ID) return;
  const data = loadRotation();
  if (data.team.length < 2) return;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const monthsIn = ['январе','феврале','марте','апреле','мае','июне','июле','августе','сентябре','октябре','ноябре','декабре'];
  const monthsOf = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
  const d = new Date(new Date().toLocaleString('en-US', { timeZone: MEETING_TZ }));
  const monthName = monthsIn[d.getMonth()];
  const monthOf = monthsOf[d.getMonth()];
  const team = data.team;

  // Анонс с обратным отсчётом
  const intro = await bot.telegram.sendMessage(GROUP_CHAT_ID,
    `🗓 <b>Новый месяц — новая ротация!</b>\n\nС 1 ${monthOf} начинаем с новой парой.\nЧерез 10 секунд узнаем кто ведёт встречи в этом месяце 👀`,
    { parse_mode: 'HTML' }
  );

  // Обратный отсчёт 10 → 1
  for (let i = 9; i >= 1; i--) {
    await sleep(1000);
    await bot.telegram.editMessageText(GROUP_CHAT_ID, intro.message_id, undefined,
      `🗓 <b>Новый месяц — новая ротация!</b>\n\nС 1 ${monthOf} начинаем с новой парой.\nЧерез ${i} ${i === 1 ? 'секунду' : i < 5 ? 'секунды' : 'секунд'} узнаем кто ведёт встречи в этом месяце 👀`,
      { parse_mode: 'HTML' }
    ).catch(() => {});
  }

  await sleep(1000);

  // Сообщение с барабаном (новое, отдельное)
  const msg = await bot.telegram.sendMessage(GROUP_CHAT_ID,
    `🎲 <b>Выбираем ведущего и секретаря на ${monthName}...</b>\n\n👤 Ведущий: <i>думаю...</i>\n📝 Секретарь: <i>думаю...</i>`,
    { parse_mode: 'HTML' }
  );
  const msgId = msg.message_id;

  // Анимация — «барабан» с случайными именами
  const spins = 7;
  for (let i = 0; i < spins; i++) {
    await sleep(i < 4 ? 500 : 700);
    const rv = team[Math.floor(Math.random() * team.length)];
    const rs = team.filter(p => p !== rv)[Math.floor(Math.random() * (team.length - 1))];
    const dots = '⏳'.repeat((i % 3) + 1);
    await bot.telegram.editMessageText(GROUP_CHAT_ID, msgId, undefined,
      `🎲 <b>Выбираем ведущего и секретаря на ${monthName}...</b>\n\n👤 Ведущий: ${dots} <i>${rv}</i>\n📝 Секретарь: ${dots} <i>${rs}</i>`,
      { parse_mode: 'HTML' }
    ).catch(() => {});
  }

  // Получаем реальное назначение
  const assignment = getMonthlyAssignment();
  if (!assignment) return;

  await sleep(900);

  // Финальное объявление
  await bot.telegram.editMessageText(GROUP_CHAT_ID, msgId, undefined,
    `🎉 <b>Ротация на ${monthName} определена!</b>\n\n👑 Ведущий: <b>${assignment.vedushiy}</b>\n📋 Секретарь: <b>${assignment.sekretar}</b>\n\nВстречи каждый вторник в 11:30 — удачи! 💪`,
    { parse_mode: 'HTML' }
  ).catch(e => console.error('animateRotation финал:', e.message));
}

// ─── Cron: автоматизация ──────────────────────────────────────────────────────
// 1. Каждый понедельник в 14:00 — создать файл на вторник + отправить анонс
cron.schedule('0 14 * * 1', async () => {
  const assignment = getMonthlyAssignment();
  if (!assignment) return;

  // Дата завтра (вторник)
  const d = new Date(new Date().toLocaleString('en-US', { timeZone: MEETING_TZ }));
  d.setDate(d.getDate() + 1);
  const months = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
  const dateStr = `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;

  // Создаём файл с tensions
  const tensions = loadSavedTensions();
  const buffer = fillProtocolTemplate({ date: dateStr, vedushiy: assignment.vedushiy, sekretar: assignment.sekretar, tensions });
  const fileName = `[${formatDateForFileName(dateStr)}] Файл встречи Дизайн Круга.docx`;
  saveToOutput(buffer, fileName);
  updateStats(assignment.vedushiy, assignment.sekretar);

  // Сохраняем мета-данные для анонса — saver.js отправит его после реального сохранения в OneDrive
  const openTensions = tensions.filter(t => t.reshili !== 'да' && t.reshili !== 'Да');
  const metaFileName = `[${formatDateForFileName(dateStr)}] meta.json`;
  const meta = {
    date: dateStr,
    vedushiy: assignment.vedushiy,
    sekretar: assignment.sekretar,
    openTensionsCount: openTensions.length
  };
  try {
    fs.writeFileSync(path.join(OUTPUT_DIR, metaFileName), JSON.stringify(meta, null, 2));
  } catch(e) { console.error('Ошибка записи meta.json:', e.message); }
}, { timezone: MEETING_TZ });

// 2. Вторник 09:00 — напоминание что файл уже в OneDrive
cron.schedule('0 9 * * 2', async () => {
  if (!NOTIFY_CHAT_ID) return;
  const assignment = getMonthlyAssignment();
  if (!assignment) return;
  const date = getTodayRussianDate();
  await bot.telegram.sendMessage(NOTIFY_CHAT_ID,
    `☀️ <b>Сегодня встреча Дизайн Круга!</b>\n\n📅 ${date} в 11:30\n👤 Ведущий: <b>${assignment.vedushiy}</b>\n📝 Секретарь: <b>${assignment.sekretar}</b>\n\n📁 Файл уже в OneDrive — не забудь заполнить до встречи`,
    { parse_mode: 'HTML' }
  ).catch(e => console.error('Cron утро встречи:', e.message));
}, { timezone: MEETING_TZ });

// 3. Напоминание секретарю в 10:30 во вторник (за час до встречи в 11:30)
cron.schedule('30 10 * * 2', async () => {
  if (!GROUP_CHAT_ID) return;
  const assignment = getMonthlyAssignment();
  if (!assignment) return;
  await bot.telegram.sendMessage(GROUP_CHAT_ID,
    `⏰ <b>Через час встреча Дизайн Круга!</b>\n\n📝 Секретарь сегодня: <b>${assignment.sekretar}</b> — не забудь подготовиться 🙌`,
    { parse_mode: 'HTML' }
  ).catch(e => console.error('Cron напоминание:', e.message));
}, { timezone: MEETING_TZ });

// 4. Ежемесячная сводка старых tensions (1-е число, 09:00)
cron.schedule('0 9 1 * *', async () => {
  const tensions = loadSavedTensions();
  const now = Date.now();
  const old = tensions.filter(t => {
    if (!t.dateAdded) return false;
    const days = (now - new Date(t.dateAdded).getTime()) / 86400000;
    return days >= 30 && t.reshili !== 'да' && t.reshili !== 'Да';
  });
  if (old.length === 0) return;
  const list = old.map((t, i) => `${i+1}. <b>${t.imya || '?'}</b>: ${t.vopros}`).join('\n');
  if (NOTIFY_CHAT_ID) await bot.telegram.sendMessage(NOTIFY_CHAT_ID,
    `📊 <b>Tensions висят 30+ дней (${old.length} шт.):</b>\n\n${list}\n\nЗакрой их в боте: /menu → ✅ Закрыть tensions`,
    { parse_mode: 'HTML' }
  ).catch(e => console.error('Cron tensions сводка:', e.message));
}, { timezone: MEETING_TZ });

// 5. Анимированная ротация 1-го числа в 09:00
cron.schedule('0 9 1 * *', async () => {
  await animateRotationInGroup().catch(e => console.error('Cron ротация:', e.message));
}, { timezone: MEETING_TZ });

bot.catch((err, ctx) => { console.error(`Ошибка Telegraf [${ctx.updateType}]:`, err); });

bot.launch().then(() => console.log('Бот запущен')).catch((err) => { console.error('Ошибка запуска:', err); process.exit(1); });

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
