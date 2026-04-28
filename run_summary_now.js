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

// Ищем любой файл встречи за сегодня, берём самый большой (заполненный)
const allFiles = fs.readdirSync(ONEDRIVE_DIR);
const matchingFiles = allFiles.filter(f => f.startsWith(`[${dateTag}]`) && f.endsWith('.docx'));
const fileName = matchingFiles.sort((a, b) => {
  return fs.statSync(path.join(ONEDRIVE_DIR, b)).size - fs.statSync(path.join(ONEDRIVE_DIR, a)).size;
})[0] || `[${dateTag}] Файл встречи Дизайн Круга.docx`;
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
        { role: 'system', content: `Ты — ассистент дизайн-команды БЦЦ. Напиши структурированное саммари встречи НА РУССКОМ, строго следуя стилю эталона.

СТРУКТУРА (только если есть данные в файле):
1. Заголовок: дата, ведущий, секретарь
2. 🏖️ Графики отпусков — кто и когда
3. ✨ Изменения в дизайн-системе
4. 👉 Задачи — ОЧЕНЬ КРАТКО, 1 строка на человека: <b>Имя</b> — над чем работает
5. 📌 Пересечения и решения — подробно
6. 📢 Новости и обучение — подробно
7. 🙏 Благодарности — подробно

ПРАВИЛА:
- Задачи: максимум 1 строка на человека, только суть, без деталей
- Пересечения, новости, благодарности — пиши полностью
- Tensions НЕ включай — добавляются отдельно
- Только HTML теги <b> и <i>, никаких <br> <ul> <div>
- Не придумывай то чего нет в файле${exampleBlock}` },
        { role: 'user', content: `Файл встречи Дизайн Круга от ${dateTag}:\n\n${rawText.substring(0, 6000)}` }
      ]
    }),
    openai.chat.completions.create({
      model: process.env.OPENAI_MODEL || 'gpt-4.1-mini',
      messages: [
        { role: 'system', content: `Ты — ассистент дизайн-команды. Извлеки tensions из протокола встречи Дизайн Круга.

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
Если tensions нет — верни [].` },
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
      if (tensionsJson.length > 0) {
        const lines = tensionsJson.map(t => {
          let s = `${t.status || '🔴'} <b>${t.imya || '?'} вызывает напряжение:</b>
${t.vopros}`;
          if (t.pochemu) s += `\n<i>Почему важно:</i> ${t.pochemu}`;
          if (t.shagi)  s += `\n<i>Возможные шаги:</i> ${t.shagi}`;
          if (t.data)   s += `\n<i>Дедлайн:</i> ${t.data}`;
          return s;
        });
        tensionsBlock = `\n\n📌 <b>Tensions встречи (${tensionsJson.length}):</b>\n\n${lines.join('\n\n')}`;
      }
    }
  } catch(e) { console.log('⚠️  Tensions не распарсились'); }

  const fullPreview = (summary + tensionsBlock)
    .replace(/<br\s*\/?>/gi, '\n')       // <br> → перенос
    .replace(/<(?!\/?(b|i|a|code|pre)\b)[^>]+>/gi, ''); // убираем все теги кроме разрешённых Telegram

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
