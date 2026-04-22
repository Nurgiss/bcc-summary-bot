require('dotenv').config();

const { Telegraf } = require('telegraf');
const OpenAI = require('openai');

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4.1-mini';

if (!TELEGRAM_BOT_TOKEN) {
  throw new Error('Не найден TELEGRAM_BOT_TOKEN в .env');
}

if (!OPENAI_API_KEY) {
  throw new Error('Не найден OPENAI_API_KEY в .env');
}

const bot = new Telegraf(TELEGRAM_BOT_TOKEN);
const openai = new OpenAI({ apiKey: OPENAI_API_KEY });

// Простое in-memory хранилище состояния
// key: `${chatId}:${userId}`
const sessions = new Map();

const QUESTIONS = [
  'Добавить что-то от себя? (например: контекст встречи, важные акценты)',
  'Ссылка на таблицу с задачами?',
  'Нужно ли что-то отдельно подсветить?'
];

function getSessionKey(ctx) {
  return `${ctx.chat.id}:${ctx.from.id}`;
}

function resetSession(ctx) {
  sessions.delete(getSessionKey(ctx));
}

function getOrCreateSession(ctx) {
  const key = getSessionKey(ctx);
  if (!sessions.has(key)) {
    sessions.set(key, {
      stage: 'waiting_raw_text',
      rawText: '',
      answers: {
        additionalContext: '',
        tasksLink: '',
        highlights: ''
      },
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
  // Поддержка формата:
  // 1: ...
  // 2: ...
  // 3: ...
  const match1 = text.match(/(?:^|\n)\s*1[\).:-]?\s*([\s\S]*?)(?=\n\s*2[\).:-]?\s*|$)/i);
  const match2 = text.match(/(?:^|\n)\s*2[\).:-]?\s*([\s\S]*?)(?=\n\s*3[\).:-]?\s*|$)/i);
  const match3 = text.match(/(?:^|\n)\s*3[\).:-]?\s*([\s\S]*)$/i);

  if (!match1 || !match2 || !match3) return null;

  return {
    additionalContext: normalizeSkip(match1[1] || ''),
    tasksLink: normalizeSkip(match2[1] || ''),
    highlights: normalizeSkip(match3[1] || '')
  };
}

function buildPrompt({ rawText, additionalContext, tasksLink, highlights }) {
  return `Ты структурируешь протокол встречи дизайнеров.

ПРАВИЛА:
— Ничего не придумывай
— Не добавляй новых фактов
— Не сокращай смысл
— Улучши читаемость
— Убери лишний шум и дубли
— Полностью ИГНОРИРУЙ задачи (tasks, планируется, to-do и т.д.)
— Не выводи задачи ни в каком виде

ВАЖНО:
— Если пользователь передал ссылку — добавь строку "Ссылка на задачи: ..."
— Если пользователь добавил комментарий — аккуратно встроить его в краткое содержание

СТРУКТУРА (строго соблюдать):

📎 Протокол встречи дизайнеров [дата]

Ссылка на таблицу: [если есть]

КРАТКОЕ СОДЕРЖАНИЕ
(2–3 предложения, живым языком)

КЛЮЧЕВЫЕ МОМЕНТЫ

🌞 Графики отпусков:
(список)

🛠 Пересечения и решения:
(список)

📌 Tension и предложения:

Актуальные:
(список)

В работе:
(список)

Решённые:
(список)

🙏 Благодарности:
(список)

ДОПОЛНИТЕЛЬНО:
— Упомяни, что добавили колонку по tension (если это есть в тексте)
— Ничего не выдумывать сверх входных данных

ВХОДНЫЕ ДАННЫЕ:

[Исходный текст]
${rawText || '—'}

[Комментарий пользователя / дополнительный контекст]
${additionalContext || '—'}

[Ссылка на таблицу с задачами]
${tasksLink || '—'}

[Что нужно отдельно подсветить]
${highlights || '—'}

Сформируй итоговый протокол строго по структуре выше.`;
}

async function generateProtocol(session) {
  const prompt = buildPrompt({
    rawText: session.rawText,
    additionalContext: session.answers.additionalContext,
    tasksLink: session.answers.tasksLink,
    highlights: session.answers.highlights
  });

  const response = await openai.responses.create({
    model: OPENAI_MODEL,
    input: [
      {
        role: 'system',
        content: 'Ты аккуратный редактор протоколов. Следуй инструкции пользователя строго.'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.2
  });

  return (response.output_text || '').trim();
}

async function sendLongMessage(ctx, text) {
  const maxLen = 3900;
  if (text.length <= maxLen) {
    await ctx.reply(text);
    return;
  }

  let start = 0;
  while (start < text.length) {
    const chunk = text.slice(start, start + maxLen);
    await ctx.reply(chunk);
    start += maxLen;
  }
}

bot.start(async (ctx) => {
  resetSession(ctx);
  getOrCreateSession(ctx);
  await ctx.reply(
    'Привет! Отправь сырой текст встречи, и я задам 3 уточняющих вопроса перед генерацией протокола.'
  );
});

bot.command('cancel', async (ctx) => {
  resetSession(ctx);
  await ctx.reply('Ок, текущий сценарий сброшен. Отправь новый сырой текст, чтобы начать заново.');
});

bot.on('text', async (ctx) => {
  const session = getOrCreateSession(ctx);
  const text = (ctx.message.text || '').trim();

  try {
    if (session.stage === 'waiting_raw_text') {
      session.rawText = text;
      session.stage = 'collecting_answers';
      session.currentQuestionIndex = 0;

      await ctx.reply(
        'Принято ✅ Теперь уточню контекст.\n\n' +
          'Можно ответить:\n' +
          '1) по шагам (я задам вопросы по очереди),\n' +
          'или 2) одним сообщением в формате:\n' +
          '1: ...\n2: ...\n3: ...\n\n' +
          `Вопрос 1/3:\n${QUESTIONS[0]}`
      );
      return;
    }

    if (session.stage === 'collecting_answers') {
      // Если пользователь прислал сразу все ответы одним сообщением
      if (session.currentQuestionIndex === 0) {
        const parsed = tryParseNumberedAnswers(text);
        if (parsed) {
          session.answers = parsed;
          session.currentQuestionIndex = 3;
        }
      }

      // Пошаговый режим
      if (session.currentQuestionIndex < 3) {
        if (session.currentQuestionIndex === 0) {
          session.answers.additionalContext = normalizeSkip(text);
          session.currentQuestionIndex = 1;
          await ctx.reply(`Вопрос 2/3:\n${QUESTIONS[1]}`);
          return;
        }

        if (session.currentQuestionIndex === 1) {
          session.answers.tasksLink = normalizeSkip(text);
          session.currentQuestionIndex = 2;
          await ctx.reply(`Вопрос 3/3:\n${QUESTIONS[2]}`);
          return;
        }

        if (session.currentQuestionIndex === 2) {
          session.answers.highlights = normalizeSkip(text);
          session.currentQuestionIndex = 3;
        }
      }

      // Все ответы собраны -> генерация
      if (session.currentQuestionIndex >= 3) {
        session.stage = 'generating';
        await ctx.reply('Готово, собираю финальный протокол...');

        const finalProtocol = await generateProtocol(session);
        if (!finalProtocol) {
          throw new Error('OpenAI вернул пустой ответ');
        }

        await sendLongMessage(ctx, finalProtocol);
        resetSession(ctx);
        return;
      }
    }

    if (session.stage === 'generating') {
      await ctx.reply('Секунду, ещё генерирую предыдущий запрос.');
    }
  } catch (error) {
    console.error('Ошибка обработки сообщения:', error);
    session.stage = 'collecting_answers';
    await ctx.reply('Произошла ошибка при обработке. Попробуй отправить ответ ещё раз или введи /cancel.');
  }
});

bot.catch((err, ctx) => {
  console.error(`Ошибка Telegraf для ${ctx.updateType}:`, err);
});

bot
  .launch()
  .then(() => {
    console.log('Бот запущен');
  })
  .catch((err) => {
    console.error('Ошибка запуска бота:', err);
    process.exit(1);
  });

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
