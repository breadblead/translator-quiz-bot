const xlsx = require('xlsx');
const path = require('path');

const medicineQuestions = [
  // ===== ЧАСТЬ 1: 1-25 =====
  { theme: 'Медицина', word: 'Scalpel', opt1: 'Шприц', opt2: 'Скальпель', opt3: 'Пинцет', opt4: 'Зажим', correct: 2 },
  { theme: 'Медицина', word: 'Emergency room', opt1: 'Операционная', opt2: 'Палата интенсивной терапии', opt3: 'Отделение неотложной помощи', opt4: 'Регистратура', correct: 3 },
  { theme: 'Медицина', word: 'Blood pressure', opt1: 'Давление крови', opt2: 'Артериальное давление', opt3: 'Пульс', opt4: 'Температура тела', correct: 2 },
  { theme: 'Медицина', word: 'Diagnosis', opt1: 'Лечение', opt2: 'Симптом', opt3: 'Рецепт', opt4: 'Диагноз', correct: 4 },
  { theme: 'Медицина', word: 'Prescription', opt1: 'Диагноз', opt2: 'Рецепт', opt3: 'Процедура', opt4: 'Анализ', correct: 2 },
  { theme: 'Медицина', word: 'Stethoscope', opt1: 'Термометр', opt2: 'Тонометр', opt3: 'Стетоскоп', opt4: 'Отоскоп', correct: 3 },
  { theme: 'Медицина', word: 'Syringe', opt1: 'Шприц', opt2: 'Игла', opt3: 'Катетер', opt4: 'Скальпель', correct: 1 },
  { theme: 'Медицина', word: 'Anesthesia', opt1: 'Антисептик', opt2: 'Анестезия', opt3: 'Анальгетик', opt4: 'Антибиотик', correct: 2 },
  { theme: 'Медицина', word: 'Cardiologist', opt1: 'Невролог', opt2: 'Кардиолог', opt3: 'Пульмонолог', opt4: 'Гастроэнтеролог', correct: 2 },
  { theme: 'Медицина', word: 'Inpatient', opt1: 'Амбулаторный пациент', opt2: 'Стационарный пациент', opt3: 'Экстренный пациент', opt4: 'Плановый пациент', correct: 2 },
  { theme: 'Медицина', word: 'Outpatient', opt1: 'Стационарный пациент', opt2: 'Амбулаторный пациент', opt3: 'Тяжёлый пациент', opt4: 'Послеоперационный пациент', correct: 2 },
  { theme: 'Медицина', word: 'Fracture', opt1: 'Ушиб', opt2: 'Растяжение', opt3: 'Перелом', opt4: 'Вывих', correct: 3 },
  { theme: 'Медицина', word: 'Concussion', opt1: 'Сотрясение мозга', opt2: 'Ушиб мозга', opt3: 'Перелом черепа', opt4: 'Кровоизлияние в мозг', correct: 1 },
  { theme: 'Медицина', word: 'Biopsy', opt1: 'Анализ крови', opt2: 'Биопсия', opt3: 'Вскрытие', opt4: 'Рентген', correct: 2 },
  { theme: 'Медицина', word: 'Chemotherapy', opt1: 'Лучевая терапия', opt2: 'Химиотерапия', opt3: 'Гормональная терапия', opt4: 'Иммунотерапия', correct: 2 },
  { theme: 'Медицина', word: 'Intensive Care Unit', opt1: 'Приёмное отделение', opt2: 'Отделение реанимации и интенсивной терапии', opt3: 'Хирургическое отделение', opt4: 'Терапевтическое отделение', correct: 2 },
  { theme: 'Медицина', word: 'Ventilator', opt1: 'Дефибриллятор', opt2: 'Кардиомонитор', opt3: 'Аппарат ИВЛ', opt4: 'Электрокардиограф', correct: 3 },
  { theme: 'Медицина', word: 'Tourniquet', opt1: 'Бинт', opt2: 'Шина', opt3: 'Жгут', opt4: 'Пластырь', correct: 3 },
  { theme: 'Медицина', word: 'Gauze', opt1: 'Марля', opt2: 'Вата', opt3: 'Бинт эластичный', opt4: 'Лейкопластырь', correct: 1 },
  { theme: 'Медицина', word: 'Suture', opt1: 'Разрез', opt2: 'Шов', opt3: 'Прокол', opt4: 'Дренаж', correct: 2 },
  { theme: 'Медицина', word: 'Inflammation', opt1: 'Инфекция', opt2: 'Воспаление', opt3: 'Кровотечение', opt4: 'Опухоль', correct: 2 },
  { theme: 'Медицина', word: 'Antibiotic', opt1: 'Антисептик', opt2: 'Противовирусное', opt3: 'Антибиотик', opt4: 'Противогрибковое', correct: 3 },
  { theme: 'Медицина', word: 'Vaccine', opt1: 'Лекарство', opt2: 'Сыворотка', opt3: 'Вакцина', opt4: 'Антидот', correct: 3 },
  { theme: 'Медицина', word: 'Dosage', opt1: 'Доза', opt2: 'Курс', opt3: 'Схема', opt4: 'Периодичность', correct: 1 },
  { theme: 'Медицина', word: 'Side effect', opt1: 'Противопоказание', opt2: 'Передозировка', opt3: 'Побочный эффект', opt4: 'Аллергическая реакция', correct: 3 },

  // ===== ЧАСТЬ 2: 26-50 =====
  { theme: 'Медицина', word: 'Malignant tumor', opt1: 'Доброкачественная опухоль', opt2: 'Злокачественная опухоль', opt3: 'Киста', opt4: 'Метастаз', correct: 2 },
  { theme: 'Медицина', word: 'Benign tumor', opt1: 'Доброкачественная опухоль', opt2: 'Злокачественная опухоль', opt3: 'Карцинома', opt4: 'Саркома', correct: 1 },
  { theme: 'Медицина', word: 'Metastasis', opt1: 'Воспаление', opt2: 'Инфекция', opt3: 'Метастаз', opt4: 'Рецидив', correct: 3 },
  { theme: 'Медицина', word: 'Remission', opt1: 'Обострение', opt2: 'Рецидив', opt3: 'Осложнение', opt4: 'Ремиссия', correct: 4 },
  { theme: 'Медицина', word: 'Relapse', opt1: 'Ремиссия', opt2: 'Рецидив', opt3: 'Выздоровление', opt4: 'Профилактика', correct: 2 },
  { theme: 'Медицина', word: 'Prognosis', opt1: 'Диагноз', opt2: 'Анамнез', opt3: 'Прогноз', opt4: 'Эпикриз', correct: 3 },
  { theme: 'Медицина', word: 'Etiology', opt1: 'Этиология', opt2: 'Патогенез', opt3: 'Симптоматика', opt4: 'Эпидемиология', correct: 1 },
  { theme: 'Медицина', word: 'Pathogenesis', opt1: 'Этиология', opt2: 'Патогенез', opt3: 'Диагностика', opt4: 'Лечение', correct: 2 },
  { theme: 'Медицина', word: 'Palliative care', opt1: 'Интенсивная терапия', opt2: 'Паллиативная помощь', opt3: 'Профилактическая медицина', opt4: 'Реабилитация', correct: 2 },
  { theme: 'Медицина', word: 'Hospice', opt1: 'Больница', opt2: 'Поликлиника', opt3: 'Хоспис', opt4: 'Санаторий', correct: 3 },
  { theme: 'Медицина', word: 'Defibrillator', opt1: 'Кардиостимулятор', opt2: 'Дефибриллятор', opt3: 'Электрокардиограф', opt4: 'Пульсоксиметр', correct: 2 },
  { theme: 'Медицина', word: 'Pacemaker', opt1: 'Дефибриллятор', opt2: 'Кардиостимулятор', opt3: 'Стетоскоп', opt4: 'Тонометр', correct: 2 },
  { theme: 'Медицина', word: 'Hemorrhage', opt1: 'Гематома', opt2: 'Тромбоз', opt3: 'Кровоизлияние', opt4: 'Ишемия', correct: 3 },
  { theme: 'Медицина', word: 'Thrombosis', opt1: 'Эмболия', opt2: 'Тромбоз', opt3: 'Аневризма', opt4: 'Варикоз', correct: 2 },
  { theme: 'Медицина', word: 'Embolism', opt1: 'Тромбоз', opt2: 'Эмболия', opt3: 'Инфаркт', opt4: 'Инсульт', correct: 2 },
  { theme: 'Медицина', word: 'Aneurysm', opt1: 'Тромб', opt2: 'Бляшка', opt3: 'Аневризма', opt4: 'Стеноз', correct: 3 },
  { theme: 'Медицина', word: 'Ischemia', opt1: 'Некроз', opt2: 'Ишемия', opt3: 'Гипоксия', opt4: 'Атрофия', correct: 2 },
  { theme: 'Медицина', word: 'Necrosis', opt1: 'Апоптоз', opt2: 'Регенерация', opt3: 'Некроз', opt4: 'Гипертрофия', correct: 3 },
  { theme: 'Медицина', word: 'Edema', opt1: 'Отёк', opt2: 'Покраснение', opt3: 'Зуд', opt4: 'Сыпь', correct: 1 },
  { theme: 'Медицина', word: 'Sepsis', opt1: 'Аллергия', opt2: 'Интоксикация', opt3: 'Сепсис', opt4: 'Шок', correct: 3 },
  { theme: 'Медицина', word: 'Anaphylaxis', opt1: 'Крапивница', opt2: 'Отёк Квинке', opt3: 'Анафилаксия', opt4: 'Поллиноз', correct: 3 },
  { theme: 'Медицина', word: 'Placebo', opt1: 'Лекарство', opt2: 'Плацебо', opt3: 'Вакцина', opt4: 'Антидот', correct: 2 },
  { theme: 'Медицина', word: 'Clinical trial', opt1: 'Доклиническое исследование', opt2: 'Клиническое исследование', opt3: 'Мета-анализ', opt4: 'Систематический обзор', correct: 2 },
  { theme: 'Медицина', word: 'Informed consent', opt1: 'Медицинская карта', opt2: 'Страховой полис', opt3: 'Информированное согласие', opt4: 'Врачебная тайна', correct: 3 },
  { theme: 'Медицина', word: 'Medical history', opt1: 'Эпикриз', opt2: 'Анамнез', opt3: 'Диагноз', opt4: 'Рецепт', correct: 2 },

  // ===== ЧАСТЬ 3: 51-75 =====
  { theme: 'Медицина', word: 'Hypertension', opt1: 'Гипотония', opt2: 'Гипертензия', opt3: 'Тахикардия', opt4: 'Аритмия', correct: 2 },
  { theme: 'Медицина', word: 'Hypotension', opt1: 'Гипертензия', opt2: 'Гипотония', opt3: 'Брадикардия', opt4: 'Стенокардия', correct: 2 },
  { theme: 'Медицина', word: 'Tachycardia', opt1: 'Брадикардия', opt2: 'Тахикардия', opt3: 'Аритмия', opt4: 'Экстрасистолия', correct: 2 },
  { theme: 'Медицина', word: 'Bradycardia', opt1: 'Тахикардия', opt2: 'Брадикардия', opt3: 'Гипертензия', opt4: 'Гипотония', correct: 2 },
  { theme: 'Медицина', word: 'Myocardial infarction', opt1: 'Инсульт', opt2: 'Инфаркт миокарда', opt3: 'Стенокардия', opt4: 'Перикардит', correct: 2 },
  { theme: 'Медицина', word: 'Stroke', opt1: 'Инфаркт', opt2: 'Инсульт', opt3: 'Эпилепсия', opt4: 'Мигрень', correct: 2 },
  { theme: 'Медицина', word: 'Epilepsy', opt1: 'Эпилепсия', opt2: 'Болезнь Паркинсона', opt3: 'Рассеянный склероз', opt4: 'Миастения', correct: 1 },
  { theme: 'Медицина', word: 'Parkinsons disease', opt1: 'Болезнь Альцгеймера', opt2: 'Болезнь Паркинсона', opt3: 'Эпилепсия', opt4: 'Деменция', correct: 2 },
  { theme: 'Медицина', word: 'Alzheimers disease', opt1: 'Болезнь Паркинсона', opt2: 'Болезнь Альцгеймера', opt3: 'Деменция', opt4: 'Шизофрения', correct: 2 },
  { theme: 'Медицина', word: 'Diabetes mellitus', opt1: 'Несахарный диабет', opt2: 'Сахарный диабет', opt3: 'Гипогликемия', opt4: 'Гипергликемия', correct: 2 },
  { theme: 'Медицина', word: 'Insulin', opt1: 'Глюкагон', opt2: 'Кортизол', opt3: 'Инсулин', opt4: 'Адреналин', correct: 3 },
  { theme: 'Медицина', word: 'Pneumonia', opt1: 'Бронхит', opt2: 'Пневмония', opt3: 'Туберкулёз', opt4: 'Астма', correct: 2 },
  { theme: 'Медицина', word: 'Tuberculosis', opt1: 'Пневмония', opt2: 'Туберкулёз', opt3: 'Бронхит', opt4: 'Эмфизема', correct: 2 },
  { theme: 'Медицина', word: 'Asthma', opt1: 'Аллергия', opt2: 'Бронхит', opt3: 'Астма', opt4: 'Пневмония', correct: 3 },
  { theme: 'Медицина', word: 'Bronchitis', opt1: 'Бронхит', opt2: 'Пневмония', opt3: 'Астма', opt4: 'Ларингит', correct: 1 },
  { theme: 'Медицина', word: 'Hepatitis', opt1: 'Панкреатит', opt2: 'Гепатит', opt3: 'Холецистит', opt4: 'Цирроз', correct: 2 },
  { theme: 'Медицина', word: 'Cirrhosis', opt1: 'Гепатит', opt2: 'Цирроз', opt3: 'Панкреатит', opt4: 'Холецистит', correct: 2 },
  { theme: 'Медицина', word: 'Pancreatitis', opt1: 'Гепатит', opt2: 'Панкреатит', opt3: 'Гастрит', opt4: 'Холецистит', correct: 2 },
  { theme: 'Медицина', word: 'Gastritis', opt1: 'Гастрит', opt2: 'Язва', opt3: 'Панкреатит', opt4: 'Колит', correct: 1 },
  { theme: 'Медицина', word: 'Ulcer', opt1: 'Гастрит', opt2: 'Язва', opt3: 'Эрозия', opt4: 'Полип', correct: 2 },
  { theme: 'Медицина', word: 'Appendicitis', opt1: 'Аппендицит', opt2: 'Перитонит', opt3: 'Холецистит', opt4: 'Дивертикулит', correct: 1 },
  { theme: 'Медицина', word: 'Peritonitis', opt1: 'Аппендицит', opt2: 'Перитонит', opt3: 'Панкреатит', opt4: 'Гастрит', correct: 2 },
  { theme: 'Медицина', word: 'Meningitis', opt1: 'Энцефалит', opt2: 'Менингит', opt3: 'Неврит', opt4: 'Радикулит', correct: 2 },
  { theme: 'Медицина', word: 'Encephalitis', opt1: 'Менингит', opt2: 'Энцефалит', opt3: 'Инсульт', opt4: 'Абсцесс', correct: 2 },
  { theme: 'Медицина', word: 'Abscess', opt1: 'Флегмона', opt2: 'Абсцесс', opt3: 'Фурункул', opt4: 'Карбункул', correct: 2 },

  // ===== ЧАСТЬ 4: 76-100 =====
  { theme: 'Медицина', word: 'Laparoscopy', opt1: 'Лапаротомия', opt2: 'Лапароскопия', opt3: 'Торакоскопия', opt4: 'Колоноскопия', correct: 2 },
  { theme: 'Медицина', word: 'Endoscopy', opt1: 'Эндоскопия', opt2: 'Колоноскопия', opt3: 'Гастроскопия', opt4: 'Бронхоскопия', correct: 1 },
  { theme: 'Медицина', word: 'Colonoscopy', opt1: 'Гастроскопия', opt2: 'Колоноскопия', opt3: 'Цистоскопия', opt4: 'Лапароскопия', correct: 2 },
  { theme: 'Медицина', word: 'Mammography', opt1: 'УЗИ молочных желёз', opt2: 'Маммография', opt3: 'Биопсия молочной железы', opt4: 'Термография', correct: 2 },
  { theme: 'Медицина', word: 'Ultrasound', opt1: 'Рентген', opt2: 'МРТ', opt3: 'Ультразвуковое исследование', opt4: 'КТ', correct: 3 },
  { theme: 'Медицина', word: 'MRI (Magnetic Resonance Imaging)', opt1: 'Компьютерная томография', opt2: 'Магнитно-резонансная томография', opt3: 'Позитронно-эмиссионная томография', opt4: 'Рентгенография', correct: 2 },
  { theme: 'Медицина', word: 'CT scan', opt1: 'МРТ', opt2: 'Компьютерная томография', opt3: 'УЗИ', opt4: 'ПЭТ', correct: 2 },
  { theme: 'Медицина', word: 'X-ray', opt1: 'УЗИ', opt2: 'Рентген', opt3: 'МРТ', opt4: 'КТ', correct: 2 },
  { theme: 'Медицина', word: 'Complete blood count', opt1: 'Биохимический анализ крови', opt2: 'Общий анализ крови', opt3: 'Коагулограмма', opt4: 'Липидограмма', correct: 2 },
  { theme: 'Медицина', word: 'Urinalysis', opt1: 'Анализ крови', opt2: 'Анализ мочи', opt3: 'Анализ кала', opt4: 'Бакпосев', correct: 2 },
  { theme: 'Медицина', word: 'Catheter', opt1: 'Зонд', opt2: 'Катетер', opt3: 'Дренаж', opt4: 'Шунт', correct: 2 },
  { theme: 'Медицина', word: 'Intravenous drip', opt1: 'Внутримышечная инъекция', opt2: 'Подкожная инъекция', opt3: 'Внутривенная капельница', opt4: 'Пероральный приём', correct: 3 },
  { theme: 'Медицина', word: 'Transplant', opt1: 'Имплантат', opt2: 'Трансплантат', opt3: 'Протез', opt4: 'Шунт', correct: 2 },
  { theme: 'Медицина', word: 'Prosthesis', opt1: 'Трансплантат', opt2: 'Протез', opt3: 'Имплантат', opt4: 'Ортез', correct: 2 },
  { theme: 'Медицина', word: 'Amputation', opt1: 'Резекция', opt2: 'Ампутация', opt3: 'Трепанация', opt4: 'Пункция', correct: 2 },
  { theme: 'Медицина', word: 'Resection', opt1: 'Ампутация', opt2: 'Резекция', opt3: 'Трансплантация', opt4: 'Имплантация', correct: 2 },
  { theme: 'Медицина', word: 'Cesarean section', opt1: 'Эпизиотомия', opt2: 'Кесарево сечение', opt3: 'Гистерэктомия', opt4: 'Лапаротомия', correct: 2 },
  { theme: 'Медицина', word: 'Prenatal care', opt1: 'Послеродовой уход', opt2: 'Дородовой уход', opt3: 'Неонатальный уход', opt4: 'Педиатрический уход', correct: 2 },
  { theme: 'Медицина', word: 'Pediatrics', opt1: 'Гериатрия', opt2: 'Педиатрия', opt3: 'Неонатология', opt4: 'Терапия', correct: 2 },
  { theme: 'Медицина', word: 'Geriatrics', opt1: 'Педиатрия', opt2: 'Гериатрия', opt3: 'Терапия', opt4: 'Неврология', correct: 2 },
  { theme: 'Медицина', word: 'Oncology', opt1: 'Онкология', opt2: 'Гематология', opt3: 'Иммунология', opt4: 'Радиология', correct: 1 },
  { theme: 'Медицина', word: 'Dermatology', opt1: 'Неврология', opt2: 'Дерматология', opt3: 'Аллергология', opt4: 'Ревматология', correct: 2 },
  { theme: 'Медицина', word: 'Ophthalmology', opt1: 'Отоларингология', opt2: 'Офтальмология', opt3: 'Стоматология', opt4: 'Неврология', correct: 2 },
  { theme: 'Медицина', word: 'Otolaryngology', opt1: 'Офтальмология', opt2: 'Отоларингология', opt3: 'Пульмонология', opt4: 'Кардиология', correct: 2 },
  { theme: 'Медицина', word: 'Rheumatology', opt1: 'Ревматология', opt2: 'Ортопедия', opt3: 'Травматология', opt4: 'Неврология', correct: 1 },
];

// Создаём массив для Excel
const excelData = medicineQuestions.map(q => ({
  'Тема': q.theme,
  'Слово/фраза для перевода': q.word,
  'Вариант 1': q.opt1,
  'Вариант 2': q.opt2,
  'Вариант 3': q.opt3,
  'Вариант 4': q.opt4,
  'Номер правильного ответа (от 1 до 4)': q.correct
}));

const ws = xlsx.utils.json_to_sheet(excelData);
const wb = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(wb, ws, 'Questions');

const filePath = path.join(__dirname, 'data', 'questions.xlsx');
xlsx.writeFile(wb, filePath);

console.log(`✅ Файл questions.xlsx создан!`);
console.log(`📊 Добавлено ${excelData.length} вопросов по Медицине`);