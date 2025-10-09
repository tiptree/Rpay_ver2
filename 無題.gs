// =======================================
// モード別設定
// =======================================
const CONFIGS = {
  rpay: {
    labelQuery: 'label:RPAY',
    sheetId: '1ozJCUSX7CSikxtKgiTKwOBe2FpTyFyCdoeaTowMsbMo',
    sheetName: 'Rpay',
    mailIdCol: 8,
    numRowsFromBottom: 50,
    filterTo: null, // フィルターなし
    sortRowCount: 50,  // 並び替えをする行の数

    fields: [
      { label: 'store_name', regex: /ご利用店舗[:：]?\s*(.+)/ },
      {
        label: 'amount',
        regex: /決済総額\s+(\d{1,3}(?:,\d{3})*)/,
        process: val => val.replace(/[¥￥,]/g, '').trim(),
      },
      {
        label: 'usage_date',
        regex: /ご利用日時\s+(\d{4}\/\d{2}\/\d{2}\(.\)\s+\d{2}:\d{2})/,
      },
      { label: 'usage', regex: /鎌田\s+(.+)様/ },
    ],

    // メインシート出力順
    outputHeaders: [
      'rpay_id',
      'input_date',
      'usage_date',     // ← 修正: 'date' → 'usage_date'
      'store_name',     // ← 修正: 'store' → 'store_name'
      'usage',          // ← 修正: 'name' → 'usage'
      'amount',
      'permanent_link',
      'mailid',
      'posision'        // ★追加
    ],

    // 家計簿出力順
    householdHeaders: [
      'rpay_id',
      'input_date',
      'date',
      'amount',
      'payment_method',
      'category_large_id',
      'category_small_id',
      'tag_id',
      'memo',
      'store',
      'month_id',
      'payment_date',
      'payment_month',
      'position_flag'
    ]
  },

  card: {
    labelQuery: 'label:カード明細',
    sheetId: '1nul58HGR_baa5v1HwLEToi2HaEK9OaMiMAkleJiu1Gw',
    sheetName: '明細T',
    mailIdCol: 11,
    numRowsFromBottom: 2000,
    filterTo: 'flhrandmini@gmail.com',
    sortRowCount: 50,  // 並び替えをする行の数

    fields_sokuho: [
      { label: 'usage_date', regex: /利用日:\s*(\d{4}\/\d{2}\/\d{2})/ },
      { label: 'usage', regex: /利用者:(.+)/ },
      {
        label: 'amount',
        regex: /利用金額:\s*(\d{1,3}(?:,\d{3})*)/,
        process: val => val.replace(/[¥￥,]/g, '').trim()
      },
      { label: 'posision', regex: /利用場所：(.+)/ }
    ],

    fields_normal: [
      { label: 'usage_date', regex: /利用日:\s*(\d{4}\/\d{2}\/\d{2})/ },
      { label: 'store_name', regex: /利用先:(.+)/ },
      { label: 'usage', regex: /利用者:(.+)/ },
      {
        label: 'amount',
        regex: /利用金額:\s*(\d{1,3}(?:,\d{3})*)/,
        process: val => val.replace(/[¥￥,]/g, '').trim()
      },
      { label: 'payment_month', regex: /支払月[:：]\s*(.+)/ },
      { label: 'posision', regex: /利用場所：(.+)/ }
    ],

    // メインシート出力順
    outputHeaders_sokuho: [
      'card_id',
      'input_date',
      'usage_date',
      'store_name',   // 空文字になる
      'usage',
      'amount',
      'payment_month', // 空文字になる
      'memo',          // 空文字
      'posision',
      'permanent_link',
      'mailid'
    ],

    outputHeaders_normal: [
      'card_id',
      'input_date',
      'usage_date',
      'store_name',
      'usage',
      'amount',
      'payment_month',
      'memo',
      'posision',
      'permanent_link',
      'mailid'
    ],

    // 家計簿出力順
    householdHeaders: [
      'card_id',
      'input_date',
      'usage_date',
      'amount',
      'payment_method',
      'category_large_id',
      'category_small_id',
      'tag_id',
      'memo',
      'store_name',
      'month_id',
      'payment_date',
      'payment_month',
      'position_flag'
    ]
  }
};

const HOUSEHOLD_SHEET = {
  sheetId: '1aPCfFScUrRhGcASCP0vaWqG9ZDKU3mi1VUZSj_lYEbk',
  sheets: {
    meisai: '明細T',
    getsudo: '月度M'
  },
  sortRowCount: 3000  // ★追加
};
