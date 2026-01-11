/**
 * Vocabulary Quiz Generator for Google Sheets
 * 어휘 학습 앱을 위한 문항 자동 생성 스크립트
 */

// 메뉴 생성
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('문항생성')
    .addItem('선택 셀 문항 생성', 'generateSelectedCell')
    .addItem('전체 문항 생성', 'generateAllItems')
    .addSeparator()
    .addItem('문항 유형 헤더 설정', 'setupSheet')
    .addSeparator()
    .addSubMenu(ui.createMenu('설정')
      .addItem('Gemini API Key 설정', 'setGeminiApiKey')
      .addItem('API Key 확인', 'checkApiKey')
      .addItem('API Key 삭제', 'deleteApiKey'))
    .addSeparator()
    .addItem('사용 안내', 'showHelp')
    .addToUi();
}

/**
 * Gemini API Key를 스크립트 속성에 저장
 */
function setGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Gemini API Key 설정',
    'Google AI Studio에서 발급받은 API Key를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();

    if (!apiKey) {
      ui.alert('API Key가 입력되지 않았습니다.');
      return;
    }

    // 스크립트 속성에 저장
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('GEMINI_API_KEY', apiKey);

    ui.alert('API Key가 성공적으로 저장되었습니다!');
  }
}

/**
 * 저장된 API Key 확인
 */
function checkApiKey() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');

  if (apiKey) {
    // API Key 일부만 표시 (보안)
    const maskedKey = apiKey.substring(0, 8) + '...' + apiKey.substring(apiKey.length - 4);
    ui.alert(`API Key가 설정되어 있습니다.\n\n키: ${maskedKey}`);
  } else {
    ui.alert('API Key가 설정되어 있지 않습니다.\n\n"설정 > Gemini API Key 설정"에서 API Key를 입력해주세요.');
  }
}

/**
 * 저장된 API Key 삭제
 */
function deleteApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'API Key 삭제',
    '저장된 API Key를 삭제하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('GEMINI_API_KEY');
    ui.alert('API Key가 삭제되었습니다.');
  }
}

/**
 * 스크립트 속성에서 API Key 가져오기
 */
function getGeminiApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('GEMINI_API_KEY');
}

/**
 * Gemini API 호출
 */
function callGeminiAPI(prompt) {
  const apiKey = getGeminiApiKey();

  if (!apiKey) {
    throw new Error('Gemini API Key가 설정되지 않았습니다. "문항생성 > 설정 > Gemini API Key 설정"에서 API Key를 입력해주세요.');
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 1024
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      const errorData = JSON.parse(responseText);
      throw new Error(`API 오류 (${responseCode}): ${errorData.error?.message || responseText}`);
    }

    const data = JSON.parse(responseText);
    return data.candidates[0].content.parts[0].text;
  } catch (error) {
    Logger.log('Gemini API Error: ' + error.message);
    throw error;
  }
}

// 설정 상수
const CONFIG = {
  VOCAB_START_COL: 1,      // A열: 선택박스
  TARGET_WORD_COL: 2,      // B열: Target Word
  POS_COL: 3,              // C열: Parts of Speech
  EXAMPLE_COL: 4,          // D열: Example Sentence
  SYNONYMS_COL: 5,         // E열: Synonyms/Antonyms
  CEFR_COL: 6,             // F열: CEFR/Grade
  KOREA_COL: 7,            // G열: Korea Curriculum
  ITEM_START_COL: 8,       // H열부터 문항 유형 시작
  HEADER_ROW: 1,           // 헤더 행
  DATA_START_ROW: 2        // 데이터 시작 행
};

/**
 * 선택한 셀에 문항 생성
 */
function generateSelectedCell() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const row = range.getRow();
  const col = range.getColumn();

  // 유효성 검사
  if (row < CONFIG.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('헤더 행이 아닌 데이터 행을 선택해주세요.');
    return;
  }

  if (col < CONFIG.ITEM_START_COL) {
    SpreadsheetApp.getUi().alert('H열 이후의 문항 유형 열을 선택해주세요.');
    return;
  }

  // 문항 유형 가져오기
  const itemType = sheet.getRange(CONFIG.HEADER_ROW, col).getValue();
  if (!itemType) {
    SpreadsheetApp.getUi().alert('선택한 열의 1행에 문항 유형이 없습니다.');
    return;
  }

  // 단어 데이터 가져오기
  const wordData = getWordData(sheet, row);
  if (!wordData.targetWord) {
    SpreadsheetApp.getUi().alert('해당 행에 Target Word가 없습니다.');
    return;
  }

  // 문항 생성
  const question = generateQuestion(itemType, wordData, sheet, row);
  range.setValue(question);

  SpreadsheetApp.getUi().alert('문항이 생성되었습니다!');
}

/**
 * 전체 빈 셀에 문항 생성 (최적화 버전)
 * - 배치 읽기/쓰기로 API 호출 최소화
 * - 실행 시간 초과 방지
 */
function generateAllItems() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < CONFIG.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('데이터가 없습니다.');
    return;
  }

  // 전체 데이터를 한 번에 읽기 (API 호출 최소화)
  const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
  const allData = dataRange.getValues();

  // 헤더 행에서 문항 유형 가져오기
  const headers = allData[0];

  // 단어 데이터 캐싱 (B~G열)
  const wordDataCache = [];
  for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    const rowIndex = row - 1; // 0-based index
    wordDataCache[row] = {
      targetWord: allData[rowIndex][CONFIG.TARGET_WORD_COL - 1],
      partsOfSpeech: allData[rowIndex][CONFIG.POS_COL - 1],
      exampleSentence: allData[rowIndex][CONFIG.EXAMPLE_COL - 1],
      synonymsAntonyms: allData[rowIndex][CONFIG.SYNONYMS_COL - 1],
      cefrGrade: allData[rowIndex][CONFIG.CEFR_COL - 1],
      koreaCurriculum: allData[rowIndex][CONFIG.KOREA_COL - 1]
    };
  }

  // 다른 단어 목록 (오답 보기용) 미리 준비
  const allWords = [];
  for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    const word = wordDataCache[row].targetWord;
    if (word) allWords.push({ row: row, word: word });
  }

  let generatedCount = 0;
  const updates = []; // 업데이트할 셀 목록

  // H열부터 마지막 열까지 순회
  for (let col = CONFIG.ITEM_START_COL; col <= lastCol; col++) {
    const itemType = headers[col - 1];
    if (!itemType) continue;

    // 각 데이터 행 순회
    for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
      const rowIndex = row - 1;
      const currentValue = allData[rowIndex][col - 1];

      // 빈 셀만 처리
      if (currentValue === '' || currentValue === null || currentValue === undefined) {
        const wordData = wordDataCache[row];
        if (wordData.targetWord) {
          // 오답 보기용 다른 단어들
          const otherWords = getOtherWordsFromCache(allWords, row, 3);
          const question = generateQuestionOptimized(itemType, wordData, otherWords);

          updates.push({
            row: row,
            col: col,
            value: question
          });
          generatedCount++;
        }
      }
    }
  }

  // 2D 배열로 한 번에 쓰기 (성능 대폭 향상)
  if (updates.length > 0) {
    // 기존 데이터 복사
    const outputData = allData.map(row => [...row]);

    // 업데이트 적용
    updates.forEach(update => {
      outputData[update.row - 1][update.col - 1] = update.value;
    });

    // H열부터 마지막 열까지만 한 번에 쓰기
    const writeRange = sheet.getRange(1, CONFIG.ITEM_START_COL, lastRow, lastCol - CONFIG.ITEM_START_COL + 1);
    const writeData = outputData.map(row => row.slice(CONFIG.ITEM_START_COL - 1));
    writeRange.setValues(writeData);
  }

  SpreadsheetApp.getUi().alert(`총 ${generatedCount}개의 문항이 생성되었습니다!`);
}

/**
 * 캐시에서 다른 단어들 가져오기 (오답 보기용)
 */
function getOtherWordsFromCache(allWords, currentRow, count) {
  const filtered = allWords.filter(item => item.row !== currentRow);
  return shuffleArray(filtered.map(item => item.word)).slice(0, count);
}

/**
 * 최적화된 문항 생성 (sheet 접근 없음)
 */
function generateQuestionOptimized(itemType, wordData, otherWords) {
  const itemTypeLower = itemType.toLowerCase().trim();

  // 빈칸 채우기
  if (itemTypeLower.includes('빈칸') || itemTypeLower.includes('fill') || itemTypeLower.includes('blank')) {
    return generateFillInBlank(wordData);
  }

  // 객관식
  if (itemTypeLower.includes('객관식') || itemTypeLower.includes('multiple') || itemTypeLower.includes('choice')) {
    return generateMultipleChoiceOptimized(wordData, otherWords);
  }

  // 유의어 찾기
  if (itemTypeLower.includes('유의어') || itemTypeLower.includes('synonym')) {
    return generateSynonymQuestionOptimized(wordData, otherWords);
  }

  // 반의어 찾기
  if (itemTypeLower.includes('반의어') || itemTypeLower.includes('antonym')) {
    return generateAntonymQuestionOptimized(wordData, otherWords);
  }

  // 품사 맞추기
  if (itemTypeLower.includes('품사') || itemTypeLower.includes('parts') || itemTypeLower.includes('pos')) {
    return generatePOSQuestion(wordData);
  }

  // 영영 풀이
  if (itemTypeLower.includes('영영') || itemTypeLower.includes('definition') || itemTypeLower.includes('meaning')) {
    return generateDefinitionQuestion(wordData);
  }

  // 문장 완성
  if (itemTypeLower.includes('문장') || itemTypeLower.includes('sentence')) {
    return generateSentenceCompletion(wordData);
  }

  // 철자
  if (itemTypeLower.includes('철자') || itemTypeLower.includes('spell')) {
    return generateSpellingQuestion(wordData);
  }

  // 문맥 추론
  if (itemTypeLower.includes('문맥') || itemTypeLower.includes('context')) {
    return generateContextQuestion(wordData);
  }

  // 단어 배열
  if (itemTypeLower.includes('배열') || itemTypeLower.includes('order') || itemTypeLower.includes('scramble')) {
    return generateWordOrderQuestion(wordData);
  }

  return generateFillInBlank(wordData);
}

/**
 * 객관식 문항 생성 (최적화)
 */
function generateMultipleChoiceOptimized(wordData, otherWords) {
  const word = wordData.targetWord;
  const sentence = wordData.exampleSentence;

  if (!sentence) {
    return `Choose the correct word: _____ (${wordData.partsOfSpeech})`;
  }

  const regex = new RegExp('\\b' + word + '\\b', 'gi');
  const blankSentence = sentence.replace(regex, '_____');

  const options = shuffleArray([word, ...otherWords]);
  const optionLabels = ['A', 'B', 'C', 'D'];
  const correctIndex = options.indexOf(word);

  let questionText = `Q: ${blankSentence}\n`;
  options.forEach((opt, idx) => {
    questionText += `${optionLabels[idx]}) ${opt}  `;
  });
  questionText += `\nAnswer: ${optionLabels[correctIndex]}`;

  return questionText;
}

/**
 * 유의어 문항 생성 (최적화)
 */
function generateSynonymQuestionOptimized(wordData, otherWords) {
  const word = wordData.targetWord;
  const synAnt = wordData.synonymsAntonyms;

  if (!synAnt || !synAnt.includes('/')) {
    return `Find a synonym for "${word}"`;
  }

  const parts = synAnt.split('/');
  const synonyms = parts[0].trim();

  if (!synonyms || synonyms === '-') {
    return `Find a synonym for "${word}" (No synonyms available)`;
  }

  const synonymList = synonyms.split(',').map(s => s.trim().replace('(similar)', '').trim());
  const correctSynonym = synonymList[0];

  const options = shuffleArray([correctSynonym, ...otherWords]);
  const optionLabels = ['A', 'B', 'C', 'D'];
  const correctIndex = options.indexOf(correctSynonym);

  let questionText = `Q: Which word is a synonym of "${word}"?\n`;
  options.forEach((opt, idx) => {
    questionText += `${optionLabels[idx]}) ${opt}  `;
  });
  questionText += `\nAnswer: ${optionLabels[correctIndex]}`;

  return questionText;
}

/**
 * 반의어 문항 생성 (최적화)
 */
function generateAntonymQuestionOptimized(wordData, otherWords) {
  const word = wordData.targetWord;
  const synAnt = wordData.synonymsAntonyms;

  if (!synAnt || !synAnt.includes('/')) {
    return `Find an antonym for "${word}"`;
  }

  const parts = synAnt.split('/');
  const antonyms = parts[1] ? parts[1].trim() : '';

  if (!antonyms || antonyms === '-') {
    return `Find an antonym for "${word}" (No antonyms available)`;
  }

  const antonymList = antonyms.split(',').map(s => s.trim());
  const correctAntonym = antonymList[0];

  const options = shuffleArray([correctAntonym, ...otherWords]);
  const optionLabels = ['A', 'B', 'C', 'D'];
  const correctIndex = options.indexOf(correctAntonym);

  let questionText = `Q: Which word is an antonym of "${word}"?\n`;
  options.forEach((opt, idx) => {
    questionText += `${optionLabels[idx]}) ${opt}  `;
  });
  questionText += `\nAnswer: ${optionLabels[correctIndex]}`;

  return questionText;
}

/**
 * 단어 데이터 가져오기
 */
function getWordData(sheet, row) {
  return {
    targetWord: sheet.getRange(row, CONFIG.TARGET_WORD_COL).getValue(),
    partsOfSpeech: sheet.getRange(row, CONFIG.POS_COL).getValue(),
    exampleSentence: sheet.getRange(row, CONFIG.EXAMPLE_COL).getValue(),
    synonymsAntonyms: sheet.getRange(row, CONFIG.SYNONYMS_COL).getValue(),
    cefrGrade: sheet.getRange(row, CONFIG.CEFR_COL).getValue(),
    koreaCurriculum: sheet.getRange(row, CONFIG.KOREA_COL).getValue()
  };
}

/**
 * 다른 단어들 가져오기 (오답 보기용)
 */
function getOtherWords(sheet, currentRow, count) {
  const lastRow = sheet.getLastRow();
  const allWords = [];

  for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    if (row !== currentRow) {
      const word = sheet.getRange(row, CONFIG.TARGET_WORD_COL).getValue();
      if (word) allWords.push(word);
    }
  }

  // 랜덤하게 섞어서 필요한 개수만큼 반환
  return shuffleArray(allWords).slice(0, count);
}

/**
 * 배열 랜덤 섞기
 */
function shuffleArray(array) {
  const newArray = [...array];
  for (let i = newArray.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
  }
  return newArray;
}

/**
 * 문항 유형에 따른 문항 생성
 */
function generateQuestion(itemType, wordData, sheet, row) {
  const itemTypeLower = itemType.toLowerCase().trim();

  // 빈칸 채우기 (Fill in the blank)
  if (itemTypeLower.includes('빈칸') || itemTypeLower.includes('fill') || itemTypeLower.includes('blank')) {
    return generateFillInBlank(wordData);
  }

  // 객관식 (Multiple Choice)
  if (itemTypeLower.includes('객관식') || itemTypeLower.includes('multiple') || itemTypeLower.includes('choice')) {
    return generateMultipleChoice(wordData, sheet, row);
  }

  // 유의어 찾기 (Synonym)
  if (itemTypeLower.includes('유의어') || itemTypeLower.includes('synonym')) {
    return generateSynonymQuestion(wordData, sheet, row);
  }

  // 반의어 찾기 (Antonym)
  if (itemTypeLower.includes('반의어') || itemTypeLower.includes('antonym')) {
    return generateAntonymQuestion(wordData, sheet, row);
  }

  // 품사 맞추기 (Parts of Speech)
  if (itemTypeLower.includes('품사') || itemTypeLower.includes('parts') || itemTypeLower.includes('pos')) {
    return generatePOSQuestion(wordData);
  }

  // 영영 풀이 (Definition)
  if (itemTypeLower.includes('영영') || itemTypeLower.includes('definition') || itemTypeLower.includes('meaning')) {
    return generateDefinitionQuestion(wordData);
  }

  // 문장 완성 (Sentence Completion)
  if (itemTypeLower.includes('문장') || itemTypeLower.includes('sentence')) {
    return generateSentenceCompletion(wordData);
  }

  // 철자 (Spelling)
  if (itemTypeLower.includes('철자') || itemTypeLower.includes('spell')) {
    return generateSpellingQuestion(wordData);
  }

  // 문맥 추론 (Context Clues)
  if (itemTypeLower.includes('문맥') || itemTypeLower.includes('context')) {
    return generateContextQuestion(wordData);
  }

  // 단어 배열 (Word Order)
  if (itemTypeLower.includes('배열') || itemTypeLower.includes('order') || itemTypeLower.includes('scramble')) {
    return generateWordOrderQuestion(wordData);
  }

  // 기본: 빈칸 채우기
  return generateFillInBlank(wordData);
}

/**
 * 빈칸 채우기 문항 생성
 */
function generateFillInBlank(wordData) {
  const word = wordData.targetWord;
  const sentence = wordData.exampleSentence;

  if (!sentence) {
    return `Fill in the blank: _____ (${wordData.partsOfSpeech})`;
  }

  // 문장에서 단어를 빈칸으로 대체
  const regex = new RegExp('\\b' + word + '\\b', 'gi');
  const blankSentence = sentence.replace(regex, '_____');

  return `Q: ${blankSentence}\nA: ${word}`;
}

/**
 * 객관식 문항 생성
 */
function generateMultipleChoice(wordData, sheet, row) {
  const word = wordData.targetWord;
  const sentence = wordData.exampleSentence;
  const otherWords = getOtherWords(sheet, row, 3);

  if (!sentence) {
    return `Choose the correct word: _____ (${wordData.partsOfSpeech})`;
  }

  // 문장에서 단어를 빈칸으로 대체
  const regex = new RegExp('\\b' + word + '\\b', 'gi');
  const blankSentence = sentence.replace(regex, '_____');

  // 보기 생성 (정답 포함해서 섞기)
  const options = shuffleArray([word, ...otherWords]);
  const optionLabels = ['A', 'B', 'C', 'D'];
  const correctIndex = options.indexOf(word);

  let questionText = `Q: ${blankSentence}\n`;
  options.forEach((opt, idx) => {
    questionText += `${optionLabels[idx]}) ${opt}  `;
  });
  questionText += `\nAnswer: ${optionLabels[correctIndex]}`;

  return questionText;
}

/**
 * 유의어 문항 생성
 */
function generateSynonymQuestion(wordData, sheet, row) {
  const word = wordData.targetWord;
  const synAnt = wordData.synonymsAntonyms;

  if (!synAnt || !synAnt.includes('/')) {
    return `Find a synonym for "${word}"`;
  }

  const parts = synAnt.split('/');
  const synonyms = parts[0].trim();

  if (!synonyms || synonyms === '-') {
    return `Find a synonym for "${word}" (No synonyms available)`;
  }

  const synonymList = synonyms.split(',').map(s => s.trim().replace('(similar)', '').trim());
  const correctSynonym = synonymList[0];

  const otherWords = getOtherWords(sheet, row, 3);
  const options = shuffleArray([correctSynonym, ...otherWords]);
  const optionLabels = ['A', 'B', 'C', 'D'];
  const correctIndex = options.indexOf(correctSynonym);

  let questionText = `Q: Which word is a synonym of "${word}"?\n`;
  options.forEach((opt, idx) => {
    questionText += `${optionLabels[idx]}) ${opt}  `;
  });
  questionText += `\nAnswer: ${optionLabels[correctIndex]}`;

  return questionText;
}

/**
 * 반의어 문항 생성
 */
function generateAntonymQuestion(wordData, sheet, row) {
  const word = wordData.targetWord;
  const synAnt = wordData.synonymsAntonyms;

  if (!synAnt || !synAnt.includes('/')) {
    return `Find an antonym for "${word}"`;
  }

  const parts = synAnt.split('/');
  const antonyms = parts[1] ? parts[1].trim() : '';

  if (!antonyms || antonyms === '-') {
    return `Find an antonym for "${word}" (No antonyms available)`;
  }

  const antonymList = antonyms.split(',').map(s => s.trim());
  const correctAntonym = antonymList[0];

  const otherWords = getOtherWords(sheet, row, 3);
  const options = shuffleArray([correctAntonym, ...otherWords]);
  const optionLabels = ['A', 'B', 'C', 'D'];
  const correctIndex = options.indexOf(correctAntonym);

  let questionText = `Q: Which word is an antonym of "${word}"?\n`;
  options.forEach((opt, idx) => {
    questionText += `${optionLabels[idx]}) ${opt}  `;
  });
  questionText += `\nAnswer: ${optionLabels[correctIndex]}`;

  return questionText;
}

/**
 * 품사 문항 생성
 */
function generatePOSQuestion(wordData) {
  const word = wordData.targetWord;
  const pos = wordData.partsOfSpeech;

  const allPOS = ['noun', 'verb', 'adjective', 'adverb', 'pronoun', 'preposition', 'conjunction', 'article'];
  const otherPOS = allPOS.filter(p => p.toLowerCase() !== pos.toLowerCase()).slice(0, 3);

  const options = shuffleArray([pos, ...otherPOS]);
  const optionLabels = ['A', 'B', 'C', 'D'];
  const correctIndex = options.findIndex(o => o.toLowerCase() === pos.toLowerCase());

  let questionText = `Q: What part of speech is "${word}"?\n`;
  options.forEach((opt, idx) => {
    questionText += `${optionLabels[idx]}) ${opt}  `;
  });
  questionText += `\nAnswer: ${optionLabels[correctIndex]}`;

  return questionText;
}

/**
 * 영영 풀이 문항 생성
 */
function generateDefinitionQuestion(wordData) {
  const word = wordData.targetWord;
  const pos = wordData.partsOfSpeech;
  const example = wordData.exampleSentence;

  // 간단한 정의 생성 (실제로는 사전 API 연동 권장)
  const definitions = {
    'the': 'Used to refer to a specific person or thing',
    'and': 'Used to connect words or phrases',
    'I': 'Used by a speaker to refer to himself or herself',
    'to': 'Expressing direction or location',
    'of': 'Expressing the relationship between a part and a whole',
    'a': 'Used when referring to someone or something for the first time',
    'you': 'Used to refer to the person being addressed',
    'it': 'Used to refer to a thing previously mentioned',
    'that': 'Used to identify a specific person or thing',
    'in': 'Expressing the situation of being enclosed or surrounded',
    'he': 'Used to refer to a male person or animal',
    'for': 'Having the purpose of',
    'they': 'Used to refer to people or things previously mentioned',
    'on': 'Physically in contact with and supported by a surface',
    'but': 'Used to introduce something contrasting',
    'we': 'Used by a speaker to refer to himself and others',
    'have': 'Possess, own, or hold',
    'know': 'Be aware of through observation or information',
    'with': 'Accompanied by another person or thing'
  };

  const definition = definitions[word.toLowerCase()] || `A ${pos} used in sentences like: "${example}"`;

  return `Q: What word matches this definition?\n"${definition}"\nA: ${word}`;
}

/**
 * 문장 완성 문항 생성
 */
function generateSentenceCompletion(wordData) {
  const word = wordData.targetWord;
  const sentence = wordData.exampleSentence;

  if (!sentence) {
    return `Complete the sentence using "${word}": _____`;
  }

  // 문장의 일부를 빈칸으로
  const words = sentence.split(' ');
  const targetIndex = words.findIndex(w => w.toLowerCase().includes(word.toLowerCase()));

  if (targetIndex !== -1) {
    words[targetIndex] = '_____';
    const blankSentence = words.join(' ');
    return `Q: Complete the sentence:\n${blankSentence}\nA: ${word}`;
  }

  return `Q: Use "${word}" to complete a sentence.\nExample: ${sentence}`;
}

/**
 * 철자 문항 생성
 */
function generateSpellingQuestion(wordData) {
  const word = wordData.targetWord;

  // 단어 철자를 섞어서 제시
  const scrambled = shuffleArray(word.split('')).join('');

  return `Q: Unscramble the letters to form a word:\n${scrambled.toUpperCase()}\nA: ${word}`;
}

/**
 * 문맥 추론 문항 생성
 */
function generateContextQuestion(wordData) {
  const word = wordData.targetWord;
  const sentence = wordData.exampleSentence;
  const pos = wordData.partsOfSpeech;

  if (!sentence) {
    return `Q: Based on context, what ${pos} fits here?\nA: ${word}`;
  }

  const regex = new RegExp('\\b' + word + '\\b', 'gi');
  const blankSentence = sentence.replace(regex, '_____');

  return `Q: Read the sentence and guess the missing word:\n"${blankSentence}"\nHint: It's a ${pos}.\nA: ${word}`;
}

/**
 * 단어 배열 문항 생성
 */
function generateWordOrderQuestion(wordData) {
  const sentence = wordData.exampleSentence;

  if (!sentence) {
    return `Arrange words to form a sentence with "${wordData.targetWord}"`;
  }

  // 문장의 단어들을 섞어서 제시
  const words = sentence.replace(/[.,!?]/g, '').split(' ');
  const scrambled = shuffleArray(words);

  return `Q: Arrange the words to form a correct sentence:\n${scrambled.join(' / ')}\nA: ${sentence}`;
}

/**
 * 사용 안내 표시
 */
function showHelp() {
  const helpText = `
📚 어휘 문항 생성기 사용 안내

[테이블 구조]
A열: 선택박스
B열: Target Word (단어)
C열: Parts of Speech (품사)
D열: Example Sentence (예문)
E열: Synonyms/Antonyms (유의어/반의어)
F열: CEFR/Grade (난이도)
G열: Korea Curriculum (교육과정)
H~Q열: 문항 유형별 생성 열 (10개)

[문항 유형 (H~Q열)]
H: 영한번역 빈칸채우기
I: 객관식
J: 유의어찾기
K: 반의어찾기
L: 품사맞추기
M: 영영풀이
N: 문장완성
O: 철자맞추기
P: 문맥추론
Q: 단어배열

[사용 방법]
1. '문항생성 > 문항 유형 헤더 설정' 클릭
2. 빈 셀 클릭 후 '선택 셀 문항 생성'
   또는 '전체 문항 생성' 클릭
  `;

  SpreadsheetApp.getUi().alert(helpText);
}
