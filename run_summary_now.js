// Запускает generateMeetingSummary прямо сейчас
require('dotenv').config();
const path = require('path');
const fs = require('fs');
const { execSync } = require('child_process');
const mammoth = require('mammoth');
const OpenAI = require('openai');

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
const ONEDRIVE_DIR = path.join(process.env.HOME, 'Library/CloudStorage/OneDrive-BCC/Meetings/Дизайн Круга');
const VPS_HOST = 'ubuntu@82.115.43.205';
const VPS_PASS = process.env.VPS_PASSWORD || 'Lvzfwhdf5inidiynitz@';
const BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const NOTIFY_CHAT_ID = process.env.NOTIFY_CHAT_ID;
const PENDING_SUMMARY_FILE = path.join(__dirname, 'pending_summary.json');

// Используем сегодняшнюю дату
const now = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Almaty' }));
const dd = String(now.getDate()).padStart(2, '0');
const mm = String(now.getMonth() + 1).padStart(2, '0');
const yyyy = now.getFullYear();
const dateTag = `${dd}.${mm}.${yyyy}`;
const fileName = `[${dateTag}] Файл встречи Дизайн Круга.docx`;
const filePath = path.join(ONEDRIVE_DIR, fileName);

console.log(`📅 Дата: ${dateTag}`);
console.log(`📄 Файл: ${filePath}`);
console.log(`✅ Существует: ${fs.existsSync(filePath)}\n`);

(async () => {
  if (!fs.existsSync(filePath)) { console.log('❌ Файл не найден'); process.exit(1); }

  const result = await mammoth.extractRawText({ path: filePath });
  const rawText = result.value.trim();
  console.log(`📝 Текст: ${rawText.length} символов\n`);

  if (rawText.length < 200) { console.log('⚠️  Файл почти пустой'); process.exit(1); }

  // Эталон с VPS
  let exampleText = '';
  try {
    exampleText = execSync(
      `sshpass -p '${VPS_PASS}' ssh -o StrictHostKeyChecking=no ${VPS_HOST} 'cat /home/ubuntu/bcc-bot/example_protocol.txt 2>/dev/null || echo ""'`,
      { encoding: 'utf8', timeout: 10000 }
    ).trim();
    console.log(exampleText ? `📚 Эталон загружен (${exampleText.length} симв.)` : 'ℹ️  Эталон не найден');
  } catch(e) { console.log('ℹ️  Эталон недоступен'); }

  const exampleBlock = exampleText ? `\n\nЭТАЛОН СТИЛЯ И СТРУКТУРЫ:\n${exampleText}` : '';

  console.log('\n⏳ Генерирую саммари + tensions параллельно...\n');

  const [summaryResp, tensionsResp] = await Promise.all([
    openai.chat.completions.create({
      model: process.env.OPENAI_MODEL || 'gpt-4.1-mini',
      messages: [
        { role: 'system', content: `Ты — ассистент дизайн-команды. Напиши краткое и чёткое summary встречи на русском языке.\nСтрого следуй стилю, структуре и тону эталона ниже.\nИспользуй HTML теги (<b>, <i>). Максимум 800 символов.${exampleBlock}` },
        { role: 'user', content: `Файл встречи Дизайн Круга от ${dateTag}:\n\n${rawText.substring(0, 6000)}` }
      ]
    }),
    openai.chat.completions.create({
      model: process.env.OPENAI_MODEL || 'gpt-4.1-mini',
      messages: [
        { role: 'system', content: `Извлеки tensions из протокола. Верни ТОЛЬКО валидный JSON-массив:\n[{"imya":"Имя кто поднял (или пустая строка)","vopros":"Что за вопрос/проблема","pochemu":"Почему важно — причина и влияние","shagi":"Предложенные шаги или кто поможет (или пустая строка)","reshili":"нет","data":"Дата/дедлайн если есть (или пустая строка)","status":"🔴 или 🟡"}]\nЕсли tensions нет — верни []. 🔴 = критично/новое, 🟡 = в процессе.` },
        { role: 'user', content: `Протокол встречи от ${dateTag}:\n\n${rawText.substring(0, 6000)}` }
      ]
    })
  ]);

  const summary = summaryResp.choices[0].message.content.trim();

  let tensionsBlock = '';
  let tensionsJson = [];
  try {
    const raw = tensionsResp.choices[0].message.content.trim();
    const jsonMatch = raw.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      tensionsJson = JSON.parse(jsonMatch[0]).map(t => ({ ...t, dateAdded: new Date().toISOString() }));
      if (tensionsJson.length > 0) {
        const lines = tensionsJson.map(t => `${t.status || '🔴'} <b>${t.imya || '?'}</b>: ${t.vopros}`);
        tensionsBlock = `\n\n📌 <b>Tensions встречи (${tensionsJson.length}):</b>\n${lines.join('\n')}`;
      }
    }
  } catch(e) { console.log('⚠️  Tensions не распарсились'); }

  const fullPreview = summary + tensionsBlock;

  console.log('─── САММАРИ ───────────────────────────────');
  console.log(fullPreview.replace(/<[^>]+>/g, ''));
  console.log('───────────────────────────────────────────\n');

  // Сохраняем pending
  fs.writeFileSync(PENDING_SUMMARY_FILE, JSON.stringify({
    summary: fullPreview, summaryOnly: summary,
    tensionsBlock, tensionsJson,
    rawText: rawText.substring(0, 6000), dateTag, exampleBlock
  }, null, 2));
  console.log('✅ pending_summary.json сохранён\n');

  // Отправляем превью в личку
  const previewButtons = { inline_keyboard: [
    [{ text: '✅ Отправить в группу', callback_data: 'send_summary_to_group' }],
    [{ text: '➕ Добавить к итогам', callback_data: 'add_to_summary' }],
    [{ text: '❌ Отменить', callback_data: 'cancel_summary' }]
  ]};
  const payload = JSON.stringify({ chat_id: NOTIFY_CHAT_ID, text: `👁 <b>Превью саммари встречи ${dateTag}:</b>\n\n${fullPreview}`, parse_mode: 'HTML', disable_web_page_preview: true, reply_markup: previewButtons });
  const https = require('https');
  const req = https.request(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) } }, res => {
    let b = ''; res.on('data', d => b += d);
    res.on('end', () => { const r = JSON.parse(b); console.log(r.ok ? '📨 Отправлено в личку!' : '❌ ' + r.description); });
  });
  req.write(payload); req.end();
})().catch(e => console.error('❌', e.message));
