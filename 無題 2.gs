//✅ ② updateIdMap()：IDマップ更新
function updateIdMap(rowObj) {
  const ss = SpreadsheetApp.openById(HOUSEHOLD_SHEET.sheetId);
  const sheet = ss.getSheetByName("id_map"); // ← id_map シート名が仮名です

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const usageKey = `${rowObj.usage_date}_${rowObj.amount}_${rowObj.usege}`;
  const idCol = headers.indexOf("id明細");

  let updated = false;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const existingKey = `${row[headers.indexOf("usage_date")]}_${row[headers.indexOf("amount")]}_${row[headers.indexOf("usege")]}`;

    if (existingKey === usageKey) {
      if (rowObj.rpay_id) row[headers.indexOf("rpayid")] = rowObj.rpay_id;
      if (rowObj.card_id && rowObj.posision == 1) row[headers.indexOf("card_sokuho_id")] = rowObj.card_id;
      if (rowObj.card_id && rowObj.posision == 0) row[headers.indexOf("card_kettei_id")] = rowObj.card_id;
      sheet.getRange(i + 2, 1, 1, headers.length).setValues([row]);
      updated = true;
      break;
    }
  }

  if (!updated) {
    // 新しい行として追加
    const newRow = headers.map(h => {
      if (h === 'id明細') return rowObj.id || '';
      if (h === 'rpayid') return rowObj.rpay_id || '';
      if (h === 'card_sokuho_id') return (rowObj.posision == 1 ? rowObj.card_id : '') || '';
      if (h === 'card_kettei_id') return (rowObj.posision == 0 ? rowObj.card_id : '') || '';
      if (h === 'usage_date') return rowObj.usage_date || '';
      if (h === 'amount') return rowObj.amount || '';
      if (h === 'usege') return rowObj.usege || '';
      return '';
    });
    sheet.appendRow(newRow);
  }
}

//✅ ① mergeHouseholdRow()：汎用上書き関数
function mergeHouseholdRow(newRowObj) {
  const ss = SpreadsheetApp.openById(HOUSEHOLD_SHEET.sheetId);
  const sheet = ss.getSheetByName(HOUSEHOLD_SHEET.sheets.meisai);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const usageKey = `${newRowObj.usage_date}_${newRowObj.amount}_${newRowObj.usege}`;
  const posCol = headers.indexOf("posision");
  const storeCol = headers.indexOf("store");

  let updated = false;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const existingKey = `${row[headers.indexOf("usage_date")]}_${row[headers.indexOf("amount")]}_${row[headers.indexOf("usege")]}`;

    if (existingKey === usageKey) {
      const existingPos = Number(row[posCol]);
      const newPos = Number(newRowObj.posision);

      if (newPos < existingPos) {
        // 優先度が高いので上書き（部分的）
        row[headers.indexOf("input_date")] = newRowObj.input_date;
        row[posCol] = newRowObj.posision;

        const storeNameCol = headers.indexOf("store");
        if (!row[storeNameCol] && newRowObj.store) {
          row[storeNameCol] = newRowObj.store;
        }

        sheet.getRange(i + 2, 1, 1, headers.length).setValues([row]);
        updated = true;
        break;
      } else {
        // 優先度が低いのでスキップ
        updated = true;
        break;
      }
    }
  }

  if (!updated) {
    // 追加行として追記
    const newRow = headers.map(h => newRowObj[h] || '');
    sheet.appendRow(newRow);
  }

  updateIdMap(newRowObj);
}

//✅ ③ writeHouseholdData()：rpay・card共通で呼ぶ関数
function writeHouseholdData(householdRowObjects) {
  if (!householdRowObjects || householdRowObjects.length === 0) return;

  householdRowObjects.forEach(row => {
    mergeHouseholdRow(row);
  });

  console.log(`家計簿シートに ${householdRowObjects.length} 件のデータをマージ処理しました。`);
}
