const xlsx = require('xlsx');
const path = require('path');


function loadQuestions() {
  const filePath = path.join(__dirname, '..', 'data', 'questions.xlsx');
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

 
  const rawData = xlsx.utils.sheet_to_json(sheet);

  
  const questions = {
    'Медицина': [],
    'Инженерия': [],
    'Строительство': []
  };

  for (const row of rawData) {
    const theme = row['Тема'];
    const question = {
      word: row['Слово/фраза для перевода'],
      options: [
        row['Вариант 1'],
        row['Вариант 2'],
        row['Вариант 3'],
        row['Вариант 4']
      ],
      correct: row['Номер правильного ответа (от 1 до 4)'] - 1 
    };

    if (questions[theme]) {
      questions[theme].push(question);
    }
  }

  return questions;
}

module.exports = { loadQuestions };