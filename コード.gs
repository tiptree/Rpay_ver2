function modeChange() {
  main('card');
}

// =======================================
// メイン関数
// =======================================
function main(mode) {
  const config = CONFIGS[mode];
  if (!config) {
    console.error(`不明なモードです：${mode}`);
    return;
  }

  const threads = searchContactMail(config.labelQuery);
  if (threads.length === 0) {
    console.log("対象メールが来ていません");
    return;
  }

  const existingIds = getLastNRows(config.sheetId, config.sheetName, config.numRowsFromBottom)
    .map(row => row[config.mailIdCol - 1]);

  const newMails = getThreadsData(threads, existingIds, mode);
  if (newMails.length === 0) {
    console.log("記入されています");
    return;
  }

  const newRowObjects = [];
  const householdRowObjects = [];

  newMails.forEach(mail => {
    const [uniqueid, inputDate, sendDate, body, permalink, subject, id] = mail;

    if (mode === 'card') {
      const isSokuho = subject.includes("速報版");
      const fieldDefs = isSokuho ? config.fields_sokuho : config.fields_normal;
      const fields = extractFields(body, fieldDefs);

      // カード用メインシート（オブジェクト）
      const mainRow = {
        card_id: uniqueid,
        input_date: inputDate,
        usage_date: fields.usage_date || '',
        store_name: isSokuho ? '' : (fields.store_name || ''),
        usege: fields.usege || '',
        amount: fields.amount || '',
        payment_month: isSokuho ? '' : (fields.payment_month || ''),
        memo: '',
        posision: isSokuho ? 1 : 0,
        permanent_link: permalink,
        mailid: id
      };
      newRowObjects.push(mainRow);

      // カード用家計簿シート（オブジェクト）
      const householdRow = {
        id: uniqueid,
        input_date: inputDate,
        usage_date: fields.usage_date || '',
        amount: fields.amount || '',
        payment_method: "CARD",
        category_large: "",
        category_small: "",
        tag: "",
        memo: "",
        store: isSokuho ? '' : (fields.store_name || ''),
        month_id: "",
        payment_date: "",
        payment_month_id: "",
        posision: isSokuho ? 1 : 0
      };
      householdRowObjects.push(householdRow);

    } else if (mode === 'rpay') {
      const fields = extractFields(body, config.fields);

      // RPay用メインシート（オブジェクト）
      const mainRow = {
        rpay_id: uniqueid,
        input_date: inputDate,
        usage_date: fields.date || '',
        store_name: fields.store || '',
        usege: fields.name || fields.cardType || '',
        amount: fields.amount || '',
        permanent_link: permalink,
        mailid: id
      };
      newRowObjects.push(mainRow);

      // RPay用家計簿シート（オブジェクト）
      const householdRow = {
        id: uniqueid,
        input_date: inputDate,
        usage_date: fields.date || '',
        amount: fields.amount || '',
        payment_method: "RPay",
        category_large: "",
        category_small: "",
        tag: "",
        memo: "",
        store: fields.store || '',
        month_id: "",
        payment_date: "",
        payment_month_id: "",
        posision: 2
      };
      householdRowObjects.push(householdRow);
    }
  });

  // メインシートへ追記（オブジェクト → 配列変換）
  if (newRowObjects.length > 0) {
    const rowArray = objectsToArrays(newRowObjects, config.headers);
    const rows = rowArray.slice(1);  // ここでヘッダー削除
    appendRowsToSheet(config.sheetId, config.sheetName, rows);
    console.log(`${newRowObjects.length} 件の ${mode} メールを書き込みました。`);
  }

  // 家計簿シートへ追記（共通）
  if (householdRowObjects.length > 0) {
    const rowArray = objectsToArrays(householdRowObjects, HOUSEHOLD_SHEET.headers);
    const rows = rowArray.slice(1);  // ここでヘッダー削除
    appendRowsToSheet(HOUSEHOLD_SHEET.sheetId, HOUSEHOLD_SHEET.sheets.meisai, rows);
    console.log(`家計簿に ${householdRowObjects.length} 件のデータを追加しました。`);
  }
}

// =======================================
// Gmail 検索
// =======================================
function searchContactMail(SEARCH_WORD) {
  return GmailApp.search(SEARCH_WORD, 0, 10);
}

// =======================================
// Gmailスレッドデータ抽出
// =======================================
function getThreadsData(threads, existingIds, mode) {
  const values = [];

  //console.log(`getThreadsData: スレッド数 = ${threads.length}`);

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const id = message.getId();
      if (existingIds.includes(id)) return;

      // card モードでは宛先フィルタリング
      if (mode === 'card') {
        const toAddress = message.getTo();
        if (!toAddress.includes("flhrandmini@gmail.com")) return;
      }
  //console.log(`メッセージID: ${id}, 件名: ${message.getSubject()}, 宛先: ${message.getTo()}`);
      const uniqueid = padStartWith0(Math.random().toString(16).slice(7), 8);
      const inputdate = new Date();
      const sendDate = message.getDate();
      const body = message.getPlainBody();
      const permalink = message.getThread().getPermalink();
      const subject = message.getSubject();

      values.push([uniqueid, inputdate, sendDate, body, permalink, subject, id]);
    });
  });
  return values;
}

// =======================================
// 最後のN行取得（重複チェック用）
// =======================================
function getLastNRows(spreadsheetId, sheetName, numRowsFromBottom) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const startRow = Math.max(lastRow - numRowsFromBottom + 1, 1);
  const numRows = lastRow - startRow + 1;
  return sheet.getRange(startRow, 1, numRows, lastCol).getValues();
}

// =======================================
// シートへ追記
// =======================================
function appendRowsToSheet(spreadsheetId, sheetName, rows) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

// =======================================
// メール本文からフィールド抽出
// =======================================
function extractFields(body, fieldDefs) {
  const result = {};
  fieldDefs.forEach(field => {
    const match = body.match(field.regex);
    let value = match ? match[1].trim() : '不明';
    if (field.process) {
      value = field.process(value);
    }
    result[field.label] = value;
  });
  return result;
}

// =======================================
// ID補助関数
// =======================================
function padStartWith0(number, length) {
  return number.toString().padStart(length, '0');
}

/**
 * 2次元配列をオブジェクトの配列に変換する関数
 * @param {Array<Array>} arrays - 最初の配列がヘッダー、残りがレコードとなる配列
 * @returns {Array<Object>} - オブジェクトの配列
 */
function arraysToObjects(arrays) {
    const [header, ...records] = arrays;  // 最初の配列をヘッダー、残りはレコード

    return records.map(record =>
        record.reduce((acc, value, index) => {
            acc[header[index]] = value;
            return acc;
        }, {})
    );
}

/**
 * オブジェクトを配列に変換する関数
 * @param {object} objects - 配列に変換するオブジェクトの配列
 * @returns {Array} オブジェクトを変換した配列
 */
function objectsToArrays(objects, headers) {
  const keys = headers || Object.keys(objects[0]); // ← 明示的な順序指定が可能
  const records = objects.map(obj => keys.map(key => obj[key]));
  return [keys, ...records];
}

