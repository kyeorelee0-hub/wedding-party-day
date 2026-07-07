// ── Google Apps Script ──
// 기존 코드에 아래 doGet 함수를 추가하세요.
// 시트에 "cards" 탭을 새로 만들고, 아래 형식으로 데이터를 입력하세요:
//
//   A열: 관계       (예: 겨레 친구)
//   B열: 이름       (참고용 — 웹사이트에 표시되지 않음)
//   C열: 만난 계기  (카드 뒷면에 표시될 텍스트)
//   D열: 아이콘     (이모지, 선택사항 — 비워두면 자동 배정)
//
// 1행은 헤더로 건너뜁니다.

const SHEET_ID = '1JSFhymv3dSC_jVh1jgFobamFCLYxggbY8BU5dNeYgLI';

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'cards') {
    return getCards();
  }

  if (action === 'populate') {
    return populateCards();
  }

  if (action === 'submitQuestion') {
    return submitQuestion(e);
  }

  if (action === 'questions') {
    return getQuestions();
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function populateCards() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const rsvpSheet = ss.getSheets()[0];
    let cardsSheet = ss.getSheetByName('cards');
    if (!cardsSheet) cardsSheet = ss.insertSheet('cards');

    const data = rsvpSheet.getDataRange().getValues();
    const seen = new Set();
    const names = [];
    data.slice(1).forEach(row => {
      const name = String(row[1]).trim();
      const status = row[3];
      if (status === '참석' && name && !seen.has(name)) {
        seen.add(name);
        names.push(name);
      }
    });

    cardsSheet.clearContents();
    cardsSheet.getRange(1, 1, 1, 2).setValues([['이름', '코멘트']]);
    if (names.length > 0) {
      const rows = names.map(n => [n, '']);
      cardsSheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', count: names.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getCards() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('cards');

    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ cards: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const rows = sheet.getDataRange().getValues();
    // 첫 행은 헤더
    const cards = rows.slice(1)
      .filter(row => row[1]) // B열 코멘트가 있는 행만
      .map(row => ({
        story: row[1],
      }));

    // CORS 허용 헤더 포함
    return ContentService
      .createTextOutput(JSON.stringify({ cards }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ cards: [], error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function submitQuestion(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('questions');
    if (!sheet) {
      sheet = ss.insertSheet('questions');
      sheet.appendRow(['제출 시각', '이름', '질문']);
    }
    const name = (e && e.parameter && e.parameter.name) || '익명';
    const question = (e && e.parameter && e.parameter.question) || '';
    if (question) {
      sheet.appendRow([new Date(), name, question]);
    }
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getQuestions() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('questions');
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ questions: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const rows = sheet.getDataRange().getValues();
    const questions = rows.slice(1).filter(r => r[2]).map(r => ({
      name: r[1],
      question: r[2],
    }));
    return ContentService
      .createTextOutput(JSON.stringify({ questions }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ questions: [], error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 기존 RSVP doPost (변경 없음) ──
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['제출 시각', '이름', '겨레/인상', '참석 여부', '기타 참고사항']);
    }

    sheet.appendRow([
      new Date(),
      e.parameter.name       || '',
      e.parameter.side       || '',
      e.parameter.attendance || '',
      e.parameter.note       || '',
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
