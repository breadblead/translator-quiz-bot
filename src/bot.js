require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const {
  startSession,
  getCurrentQuestion,
  checkAnswer,
  nextQuestion,
  getResult,
  clearSession
} = require('./quizLogic');

const token = process.env.BOT_TOKEN;
const bot = new TelegramBot(token, { polling: true });


function getMainMenuKeyboard() {
  return {
    reply_markup: {
      inline_keyboard: [
        [{ text: '🏥 Медицина', callback_data: 'theme_Медицина' }],
        [{ text: '⚙️ Инженерия', callback_data: 'theme_Инженерия' }],
        [{ text: '🏗️ Строительство', callback_data: 'theme_Строительство' }]
      ]
    }
  };
}


function getAnswerKeyboard(options) {
  return {
    reply_markup: {
      inline_keyboard: options.map((opt, index) => [{
        text: opt,
        callback_data: `answer_${index}`
      }])
    }
  };
}


function getNextKeyboard() {
  return {
    reply_markup: {
      inline_keyboard: [
        [{ text: '➡️ Далее', callback_data: 'next' }]
      ]
    }
  };
}


function getFinishKeyboard() {
  return {
    reply_markup: {
      inline_keyboard: [
        [{ text: '🔄 Выбрать новую тему', callback_data: 'new_theme' }],
        [{ text: '🔁 Пройти заново', callback_data: 'retry' }]
      ]
    }
  };
}



bot.onText(/\/start/, (msg) => {
  const chatId = msg.chat.id;
  const firstName = msg.from.first_name || 'Пользователь';

  const welcomeText = 
    `👋 Привет, *${firstName}*!\n\n` +
    `Я — *Тренажёр переводчика*. Помогу тебе проверить знание профессиональных терминов.\n\n` +
    `📚 Доступные темы:\n` +
    `🏥 Медицина\n` +
    `⚙️ Инженерия\n` +
    `🏗️ Строительство\n\n` +
    `В каждой сессии — *10 случайных вопросов* с 4 вариантами ответа.\n\n` +
    `👇 Выбери тему, чтобы начать:`;

  bot.sendMessage(chatId, welcomeText, {
    parse_mode: 'Markdown',
    ...getMainMenuKeyboard()
  });
});



bot.on('callback_query', async (query) => {
  const chatId = query.message.chat.id;
  const messageId = query.message.message_id;
  const data = query.data;

 
  bot.answerCallbackQuery(query.id);

  
  if (data.startsWith('theme_')) {
    const theme = data.replace('theme_', '');
    const question = startSession(chatId, theme);

    if (!question) {
      bot.sendMessage(chatId, '❌ Вопросы по этой теме не найдены.');
      return;
    }

    const text = formatQuestionText(question);
    bot.sendMessage(chatId, text, {
      parse_mode: 'HTML',
      ...getAnswerKeyboard(question.options)
    });
    return;
  }

  
  if (data.startsWith('answer_')) {
    const answerIndex = parseInt(data.replace('answer_', ''));
    const result = checkAnswer(chatId, answerIndex);
    const question = getCurrentQuestion(chatId);

    if (!result || !question) {
      bot.sendMessage(chatId, '⚠️ Сессия не найдена. Начните заново: /start');
      return;
    }

    
    const updatedKeyboard = question.options.map((opt, index) => {
      let prefix = '';
      if (index === result.correctIndex) prefix = '🟢 ';
      else if (index === result.chosenIndex && !result.isCorrect) prefix = '🔴 ';
      return [{ text: `${prefix}${opt}`, callback_data: 'answered' }];
    });

    
    const statusText = result.isCorrect ? '✅ Правильно!' : '❌ Неправильно!';
    const updatedText = formatQuestionText(question) + `\n\n${statusText}`;

    bot.editMessageText(updatedText, {
      chat_id: chatId,
      message_id: messageId,
      parse_mode: 'HTML',
      reply_markup: { inline_keyboard: updatedKeyboard }
    });

    bot.sendMessage(chatId, '⏭ Нажмите «Далее» для следующего вопроса:', getNextKeyboard());
    return;
  }

 
  if (data === 'next') {
    
    bot.deleteMessage(chatId, messageId);

    const question = nextQuestion(chatId);

    
    if (!question) {
      const result = getResult(chatId);
      const finishText = formatResultText(result);

      bot.sendMessage(chatId, finishText, {
        parse_mode: 'HTML',
        ...getFinishKeyboard()
      });

      clearSession(chatId);
      return;
    }

    
    const text = formatQuestionText(question);
    bot.sendMessage(chatId, text, {
      parse_mode: 'HTML',
      ...getAnswerKeyboard(question.options)
    });
    return;
  }

  
  if (data === 'retry') {
    
    bot.sendMessage(chatId, '👇 Выберите тему для повторного прохождения:', getMainMenuKeyboard());
    return;
  }

  
  if (data === 'new_theme') {
    bot.sendMessage(chatId, '👇 Выберите новую тему:', getMainMenuKeyboard());
    return;
  }

  
  if (data === 'answered') {
    bot.answerCallbackQuery(query.id, { text: 'Вы уже ответили на этот вопрос', show_alert: false });
    return;
  }
});



function formatQuestionText(question) {
  return (
    `📝 <b>Вопрос ${question.questionNumber}/${question.total}</b>\n\n` +
    `🔤 Переведите:\n<b>«${question.word}»</b>\n\n` +
    `👇 Выберите вариант ответа:`
  );
}

function formatResultText(result) {
  const { score, total, theme } = result;
  const percent = Math.round((score / total) * 100);
  let emoji = '';

  if (percent === 100) emoji = '🏆';
  else if (percent >= 80) emoji = '👏';
  else if (percent >= 50) emoji = '👍';
  else emoji = '📚';

  return (
    `${emoji} <b>Тест завершён!</b>\n\n` +
    `📌 Тема: <b>${theme}</b>\n` +
    `✅ Правильных ответов: <b>${score} из ${total}</b>\n` +
    `📊 Процент: <b>${percent}%</b>\n\n` +
    `Что дальше?`
  );
}



bot.on('polling_error', (error) => {
  console.error('Polling error:', error);
});

console.log('🤖 Бот «Тренажёр переводчика» запущен...');