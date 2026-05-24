const { loadQuestions } = require('./excelReader');


const sessions = {};


function shuffle(array) {
  const arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}


function startSession(chatId, theme) {
  const allQuestions = loadQuestions();
  const themeQuestions = allQuestions[theme];

  if (!themeQuestions || themeQuestions.length === 0) {
    return null;
  }

  
  const count = Math.min(10, themeQuestions.length);
  const selected = shuffle(themeQuestions).slice(0, count);

  sessions[chatId] = {
    theme,
    questions: selected,
    currentIndex: 0,
    score: 0,
    total: count
  };

  return getCurrentQuestion(chatId);
}


function getCurrentQuestion(chatId) {
  const session = sessions[chatId];
  if (!session || session.currentIndex >= session.total) return null;

  return {
    word: session.questions[session.currentIndex].word,
    options: session.questions[session.currentIndex].options,
    questionNumber: session.currentIndex + 1,
    total: session.total
  };
}


function checkAnswer(chatId, answerIndex) {
  const session = sessions[chatId];
  if (!session) return null;

  const question = session.questions[session.currentIndex];
  const isCorrect = answerIndex === question.correct;

  if (isCorrect) {
    session.score++;
  }

  return {
    isCorrect,
    correctIndex: question.correct,
    chosenIndex: answerIndex
  };
}


function nextQuestion(chatId) {
  const session = sessions[chatId];
  if (!session) return null;

  session.currentIndex++;

  if (session.currentIndex >= session.total) {
    
    return null;
  }

  return getCurrentQuestion(chatId);
}


function getResult(chatId) {
  const session = sessions[chatId];
  if (!session) return null;

  return {
    score: session.score,
    total: session.total,
    theme: session.theme
  };
}


function clearSession(chatId) {
  delete sessions[chatId];
}

module.exports = {
  startSession,
  getCurrentQuestion,
  checkAnswer,
  nextQuestion,
  getResult,
  clearSession
};