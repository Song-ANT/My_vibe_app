/**
 * 웹앱 진입점 (GET 요청 처리)
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Purrfect Budget - 지출 관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 지출 시트 가져오기 또는 생성
 */
function getExpenseSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = '지출내역';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['ID', '날짜', '금액', '카테고리', '사용처', '결제수단', '메모']);
  }
  return sheet;
}

/**
 * [화면 1] 지출 한 건 저장
 */
function saveExpense(data) {
  try {
    if (!data.date || !/^\d{4}-\d{2}-\d{2}$/.test(data.date)) {
      return { success: false, message: '날짜 형식이 올바르지 않습니다. (예: 2026-11-04)' };
    }
    var amount = Number(data.amount);
    if (isNaN(amount) || amount < 0 || String(data.amount).trim() === '') {
      return { success: false, message: '금액은 0원 이상의 숫자로 입력해 주세요.' };
    }
    if (!data.category) {
      return { success: false, message: '카테고리를 선택해 주세요.' };
    }
    if (!data.place || !data.place.trim()) {
      return { success: false, message: '사용처를 입력해 주세요.' };
    }

    var sheet = getExpenseSheet();
    var lastRow = sheet.getLastRow();
    var newId = (lastRow <= 1) ? 1 : lastRow;
    var payment = (data.payment && data.payment.trim()) ? data.payment.trim() : '--';
    var memo = (data.memo) ? String(data.memo).trim() : '';

    sheet.appendRow([
      newId,
      data.date,
      amount,
      data.category,
      data.place.trim(),
      payment,
      memo
    ]);

    var appendedRow = sheet.getLastRow();
    sheet.getRange(appendedRow, 2).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(appendedRow, 3).setNumberFormat('#,##0');

    var dateParts = data.date.split('-');
    var successMsg = parseInt(dateParts[1], 10) + '월 ' + parseInt(dateParts[2], 10) + '일 ' + data.place.trim() + ' (' + amount.toLocaleString('ko-KR') + '원) 저장되었습니다.';

    return { success: true, message: successMsg };
  } catch (err) {
    return { success: false, message: '저장 오류: ' + err.toString() };
  }
}

/**
 * [화면 1] CSV 일괄 업로드 파싱 저장
 */
function uploadExpensesCsv(csvText) {
  try {
    if (!csvText || !csvText.trim()) {
      return { success: false, message: 'CSV 파일 내용이 비어 있습니다.' };
    }

    var lines = Utilities.parseCsv(csvText);
    if (lines.length <= 1) {
      return { success: false, message: '업로드할 데이터 행이 없습니다.' };
    }

    var sheet = getExpenseSheet();
    var lastRow = sheet.getLastRow();
    var currentId = (lastRow <= 1) ? 1 : lastRow;
    
    var rowsToAdd = [];
    var startIndex = (lines[0][0].toLowerCase().indexOf('id') !== -1 || (lines[0][1] && lines[0][1].indexOf('날짜') !== -1)) ? 1 : 0;

    for (var i = startIndex; i < lines.length; i++) {
      var row = lines[i];
      if (!row || row.length < 2) continue;

      var offset = (row.length >= 7 && !isNaN(Number(row[0]))) ? 1 : 0;
      
      var dateStr = String(row[offset] || '').trim();
      var amountVal = Number(String(row[offset + 1] || '0').replace(/[^0-9.-]+/g, ''));
      var categoryVal = String(row[offset + 2] || '기타').trim() || '기타';
      var placeVal = String(row[offset + 3] || '미지정').trim() || '미지정';
      var payVal = String(row[offset + 4] || '--').trim() || '--';
      var memoVal = String(row[offset + 5] || '').trim();

      if (!dateStr) continue;

      rowsToAdd.push([
        currentId++,
        dateStr,
        isNaN(amountVal) ? 0 : amountVal,
        categoryVal,
        placeVal,
        payVal,
        memoVal
      ]);
    }

    if (rowsToAdd.length === 0) {
      return { success: false, message: '유효한 지출 내역 행을 찾을 수 없습니다.' };
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
    return { success: true, message: rowsToAdd.length + '건의 지출 내역을 성공적으로 일괄 등록했습니다.' };
  } catch (err) {
    return { success: false, message: 'CSV 업로드 중 오류 발생: ' + err.toString() };
  }
}

/**
 * [화면 2 & 3] 전체 데이터 조회 (오류 해결 핵심 함수)
 */
function getExpenseData() {
  try {
    var sheet = getExpenseSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, list: [] };
    }

    var values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var list = [];

    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      var d = r[1];
      var dateStr = '';
      if (d instanceof Date) {
        var yyyy = d.getFullYear();
        var mm = String(d.getMonth() + 1).padStart(2, '0');
        var dd = String(d.getDate()).padStart(2, '0');
        dateStr = yyyy + '-' + mm + '-' + dd;
      } else {
        dateStr = String(d || '');
      }

      list.push({
        id: r[0],
        date: dateStr,
        amount: Number(r[2]) || 0,
        category: String(r[3] || '기타'),
        place: String(r[4] || ''),
        payment: String(r[5] || '--'),
        memo: String(r[6] || '')
      });
    }

    return { success: true, list: list };
  } catch (err) {
    return { success: false, message: '데이터 불러오기 오류: ' + err.toString() };
  }
}
