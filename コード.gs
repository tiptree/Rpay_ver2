const IS_TEST_MODE = true; // ← 本番では false


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

  const resolver = new GatudoResolver(); // ← これを追加

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

  //ここにメール上書き関数呼び出しを追加
  if (householdRowObjects.length > 0) {
    writeHouseholdData(householdRowObjects);  // ← これを使う
  }

  //ここまで


  newMails.forEach(mail => {
    const [uniqueid, inputDate, sendDate, body, permalink, subject, id] = mail;
    if (mode === 'card') {
      const isSokuho = subject.includes("速報版");
      const fieldDefs = isSokuho ? config.fields_sokuho : config.fields_normal;
      const fields = extractFields(body, fieldDefs);

      // ★ 追加：card用 usage_date を Date 化
      const usageDate = fields.usage_date
        ? new Date(fields.usage_date.replace(/\//g, '-'))
        : null;


      // カード用メインシート（オブジェクト）
      const mainRow = {
        card_id: uniqueid,
        input_date: inputDate,
        usage_date: usageDate,
        store_name: isSokuho ? '' : (fields.store_name || ''),
        usage: fields.usage || '',
        amount: fields.amount || '',
        payment_month: isSokuho ? '' : (fields.payment_month || ''),
        memo: '',
        position: isSokuho ? 1 : 0,
        permanent_link: permalink,
        mailid: id
      };
      newRowObjects.push(mainRow);

      // カード　→　家計簿シート（オブジェクト）
      const householdRow = {
        id: uniqueid,
        input_date: inputDate,
        usage_date: usageDate,
        amount: fields.amount || '',
        payment_method: "CARD",
        category_large: "",
        category_small: "",
        tag: fields.usage || '',  //→usageを入力する　自分用か家族用か
        memo: "",
        store_name: isSokuho ? '' : (fields.store_name || ''),
        month_id: resolver.resolve(usageDate),
        payment_date: "",
        payment_month_id: "",
        position: isSokuho ? 1 : 0
      };
      householdRowObjects.push(householdRow);

    } else if (mode === 'rpay') {
      const fields = extractFields(body, config.fields);

      // "2025/12/17(水) 15:32" → Date
      const usageDateStr = fields.usage_date;
      const usageDate = new Date(
        usageDateStr
          .replace(/\(.+\)/, '')  // 曜日削除
          .replace(/\//g, '-')
      );


      // RPay用メインシート（オブジェクト）
      const mainRow = {
        rpay_id: uniqueid,
        input_date: inputDate,
        usage_date: usageDate,
        store_name: fields.store_name || '',
        usage: fields.usage || fields.cardType || '',
        amount: fields.amount || '',
        permanent_link: permalink,
        mailid: id,
        position: 2
      };
      newRowObjects.push(mainRow);

      // RPay　→　家計簿シート（オブジェクト）
      const householdRow = {
        id: uniqueid,
        input_date: inputDate,
        usage_date: usageDate,
        amount: fields.amount || '',
        payment_method: "RPay",
        category_large: "",
        category_small: "",
        tag: fields.usage === "陽介" ? "本人" : fields.usage === "美智子" ? "家族" : "不明",
        memo: "",
        store_name: fields.store_name || '',
        month_id: resolver.resolve(usageDate), // ← ここで月度IDをセット,
        payment_date: "",
        payment_month_id: "",
        position: 2
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

  // すべて書き込み完了後に、並び替えを一括で実行
  sortSheet(HOUSEHOLD_SHEET.sheetId, HOUSEHOLD_SHEET.sheets.meisai, HOUSEHOLD_SHEET.sortRowCount);
  sortSheet(config.sheetId, config.sheetName, config.sortRowCount);


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
    // ★ 安全に field.process を適用
    if (field.process && value !== '不明') {
      try {
        value = field.process(value);
      } catch (e) {
        console.warn(`処理中にエラーが発生しました（${field.label}）:`, e);
      }
    }

    result[field.label] = value;
  });
  //console.log(result); // ← 追加して、GASログで確認
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

class GatudoResolver {
  constructor() {
    const ss = SpreadsheetApp.openById(HOUSEHOLD_SHEET.sheetId);
    const sheet = ss.getSheetByName(HOUSEHOLD_SHEET.sheets.getsudo);
    const values = sheet.getDataRange().getValues();

    // 1行目はヘッダーなので除外
    this.monthRanges = values.slice(1).map(row => {
      const [gatudo_id, , startDateRaw, endDateRaw] = row;

      const start = this._toDateOnly(startDateRaw);
      const end = this._toDateOnly(endDateRaw);

      return { id: gatudo_id, start, end };
    });
  }

  /**
   * Date または文字列を「日付のみ（00:00:00）」の Date に変換
   */
  _toDateOnly(value) {
    let d;
    if (value instanceof Date) {
      d = value;
    } else {
      const clean = value.replace(/\(.+\)/, '').trim(); // 曜日削除
      d = new Date(clean.replace(/\//g, '-'));
    }
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  /**
   * usageDate は Date オブジェクト（時刻あり）を受け取る
   */
  resolve(usageDate) {
    const compareDate = new Date(
      usageDate.getFullYear(),
      usageDate.getMonth(),
      usageDate.getDate()
    );

    for (const range of this.monthRanges) {
      if (compareDate >= range.start && compareDate <= range.end) {
        return range.id;
      }
    }
    return '';
  }
}



/**
 * 並び替え数
 * 
 * 
 */
function sortSheet(sheetId, sheetName, sortRowCount) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let positionCol = headers.indexOf('position') + 1;
  const usageDateCol = headers.indexOf('usage_date') + 1;

  if (usageDateCol === 0) {
    console.error(`usage_date 列が見つかりません（${sheetName}）`);
    return;
  }

  const rangeRowCount = Math.min(lastRow - 1, sortRowCount);
  if (rangeRowCount <= 0) {
    console.log(`${sheetName} に並び替え対象のデータがありません`);
    return;
  }

  const sortCriteria = [
    { column: positionCol, ascending: true },
    { column: usageDateCol, ascending: true }
  ];

  sheet.getRange(2, 1, rangeRowCount, lastCol).sort(sortCriteria);
  console.log(`${sheetName} を並び替え完了（行数: ${rangeRowCount}）`);
}

