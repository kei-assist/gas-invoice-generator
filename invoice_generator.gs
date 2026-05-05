// ============================================================
// 請求書・見積書 自動生成ツール for Google Sheets
// フリーランス・個人事業主向け
// ============================================================

const SHEET_INPUT    = '入力';
const SHEET_INVOICE  = '請求書';
const SHEET_SETTINGS = '設定';

// ============================================================
// メニュー登録
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📄 請求書ツール')
    .addItem('▶ 請求書を生成', 'generateInvoice')
    .addItem('▶ 見積書を生成', 'generateQuote')
    .addSeparator()
    .addItem('↺ 入力シートをリセット', 'resetInput')
    .addSeparator()
    .addItem('⚙ 初期セットアップ（初回のみ）', 'setupSheets')
    .addToUi();
}

// ============================================================
// 初期セットアップ
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupSettingsSheet(ss);
  _setupInputSheet(ss);
  _setupInvoiceSheet(ss);
  ss.setActiveSheet(ss.getSheetByName(SHEET_SETTINGS));
  SpreadsheetApp.getUi().alert(
    'セットアップ完了！',
    '「設定」シートに自社情報・振込先を入力してください。\n入力後は「入力」シートから請求書を作成できます。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function _setupSettingsSheet(ss) {
  let sh = ss.getSheetByName(SHEET_SETTINGS) || ss.insertSheet(SHEET_SETTINGS);
  sh.clear(); sh.clearFormats();

  const rows = [
    ['【自社情報】', ''],
    ['', ''],
    ['会社名 / 屋号',           '例）Kei Assist'],
    ['氏名',                    '例）藤田 啓一'],
    ['郵便番号',                '例）100-0001'],
    ['住所',                    '例）東京都千代田区〇〇1-2-3'],
    ['電話番号',                '例）090-0000-0000'],
    ['メールアドレス',          '例）your@email.com'],
    ['インボイス登録番号',      '例）T1234567890123'],
    ['', ''],
    ['【振込先情報】', ''],
    ['', ''],
    ['銀行名',   '例）〇〇銀行'],
    ['支店名',   '例）〇〇支店'],
    ['口座種別', '普通'],
    ['口座番号', '例）1234567'],
    ['口座名義', '例）フジタ ケイイチ'],
    ['', ''],
    ['【書類設定】', ''],
    ['', ''],
    ['消費税率（%）',            '10'],
    ['支払期限（発行日から日数）','30'],
    ['番号プレフィックス',       'INV'],
  ];

  sh.getRange(1, 1, rows.length, 2).setValues(rows);

  // タイトル行スタイル
  [[1,1],[11,1],[19,1]].forEach(([r]) => {
    sh.getRange(r, 1, 1, 2).merge()
      .setBackground('#1a73e8').setFontColor('white')
      .setFontWeight('bold').setFontSize(11);
  });

  // ラベル列
  sh.getRange(1, 1, rows.length, 1).setFontWeight('bold');
  sh.getRange(3, 2, rows.length - 2, 1).setBackground('#fffde7');

  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 320);
  sh.setFrozenRows(1);
}

function _setupInputSheet(ss) {
  let sh = ss.getSheetByName(SHEET_INPUT) || ss.insertSheet(SHEET_INPUT);
  sh.clear(); sh.clearFormats();

  // 列幅
  [40, 200, 70, 60, 110, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ===== タイトルバー =====
  sh.getRange('A1:F1').merge()
    .setValue('📋  請求書・見積書  入力フォーム')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#1a73e8').setFontColor('white')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 36);

  // ===== 書類基本情報 =====
  _label(sh, 3, 1, '書類タイプ');
  sh.getRange(3, 2).setValue('請求書').setBackground('#fffde7');
  sh.getRange(3, 2).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['請求書', '見積書'], true).build()
  );

  _label(sh, 4, 1, '発行日');
  sh.getRange(4, 2).setValue(new Date()).setNumberFormat('yyyy/MM/dd').setBackground('#fffde7');

  _label(sh, 4, 4, '支払期限');
  sh.getRange(4, 5, 1, 2).merge().setNumberFormat('yyyy/MM/dd').setBackground('#fffde7');

  // ===== 請求先 =====
  sh.getRange('A6:F6').merge()
    .setValue('▼ 請求先情報')
    .setBackground('#e8f0fe').setFontWeight('bold');

  const clientFields = [
    [7, '会社名 / 氏名', true],
    [8, '担当者名',     false],
    [9, '郵便番号',     false],
    [10,'住所',         true],
  ];
  clientFields.forEach(([r, label, wide]) => {
    _label(sh, r, 1, label);
    const range = wide
      ? sh.getRange(r, 2, 1, 5).merge()
      : sh.getRange(r, 2);
    range.setBackground('#fffde7');
  });

  _label(sh, 12, 1, '件名');
  sh.getRange(12, 2, 1, 5).merge().setBackground('#fffde7');

  _label(sh, 13, 1, '備考');
  sh.getRange(13, 2, 1, 5).merge().setBackground('#fffde7').setWrap(true);
  sh.setRowHeight(13, 60);

  // ===== 品目テーブル =====
  sh.getRange('A15:F15').merge()
    .setValue('▼ 品目・明細')
    .setBackground('#e8f0fe').setFontWeight('bold');

  const headers = ['No.', '品目・内容', '数量', '単位', '単価', '金額（自動）'];
  sh.getRange(16, 1, 1, 6).setValues([headers])
    .setBackground('#34a853').setFontColor('white')
    .setFontWeight('bold').setHorizontalAlignment('center');

  for (let i = 17; i <= 31; i++) {
    const no = i - 16;
    const bg = no % 2 === 0 ? '#f8f9fa' : '#ffffff';
    sh.getRange(i, 1).setValue(no).setHorizontalAlignment('center').setFontColor('#aaaaaa').setBackground(bg);
    sh.getRange(i, 2).setBackground(bg);
    sh.getRange(i, 3).setBackground(bg).setHorizontalAlignment('right').setNumberFormat('#,##0');
    sh.getRange(i, 4).setBackground(bg).setHorizontalAlignment('center');
    sh.getRange(i, 5).setBackground(bg).setHorizontalAlignment('right').setNumberFormat('#,##0');
    sh.getRange(i, 6)
      .setFormula(`=IFERROR(IF(B${i}="","",C${i}*E${i}),"")`)
      .setBackground('#f0f4ff').setHorizontalAlignment('right').setNumberFormat('#,##0');
  }

  // 合計行
  const totBg = '#e8f0fe';
  _totalRow(sh, 32, '小　計',         `=SUM(F17:F31)`, totBg, false);
  _totalRow(sh, 33, '消費税（10%）',   `=ROUND(F32*0.1,0)`, totBg, false);
  _totalRow(sh, 34, '合計金額（税込）',`=F32+F33`, '#fbbc04', true);

  sh.setFrozenRows(16);
}

function _setupInvoiceSheet(ss) {
  let sh = ss.getSheetByName(SHEET_INVOICE) || ss.insertSheet(SHEET_INVOICE);
  sh.clear(); sh.clearFormats();
  sh.getRange('A1').setValue('← 「入力」シートに情報を入力し、メニューから請求書/見積書を生成してください')
    .setFontColor('#aaaaaa').setFontStyle('italic');
}

// ===== ヘルパー =====
function _label(sh, row, col, text) {
  sh.getRange(row, col).setValue(text)
    .setFontWeight('bold').setBackground('#f5f5f5');
}
function _totalRow(sh, row, label, formula, bg, bold) {
  sh.getRange(row, 1, 1, 5).merge()
    .setValue(label).setHorizontalAlignment('right')
    .setFontWeight(bold ? 'bold' : 'normal')
    .setFontSize(bold ? 12 : 10).setBackground(bg);
  sh.getRange(row, 6)
    .setFormula(formula).setHorizontalAlignment('right')
    .setNumberFormat('#,##0').setBackground(bg)
    .setFontWeight(bold ? 'bold' : 'normal')
    .setFontSize(bold ? 12 : 10);
}

// ============================================================
// 設定・入力データ取得
// ============================================================
function _getSettings() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) throw new Error('設定シートがありません。「初期セットアップ」を実行してください。');

  const vals = sh.getRange('A1:B' + sh.getLastRow()).getValues();
  const map  = {};
  vals.forEach(([k, v]) => { if (k && v !== '') map[String(k).trim()] = v; });

  return {
    companyName:   map['会社名 / 屋号']           || '',
    ownerName:     map['氏名']                    || '',
    zip:           map['郵便番号']                || '',
    address:       map['住所']                    || '',
    phone:         map['電話番号']                || '',
    email:         map['メールアドレス']           || '',
    invoiceReg:    map['インボイス登録番号']        || '',
    bankName:      map['銀行名']                  || '',
    bankBranch:    map['支店名']                  || '',
    bankType:      map['口座種別']                || '普通',
    bankAccount:   map['口座番号']                || '',
    bankHolder:    map['口座名義']                || '',
    taxRate:       parseFloat(map['消費税率（%）'] || '10') / 100,
    payDays:       parseInt(map['支払期限（発行日から日数）'] || '30'),
    prefix:        map['番号プレフィックス']       || 'INV',
  };
}

function _getInputData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_INPUT);
  if (!sh) throw new Error('入力シートがありません。「初期セットアップ」を実行してください。');

  const v = (r, c) => sh.getRange(r, c).getValue();

  const items = [];
  for (let i = 17; i <= 31; i++) {
    const name = v(i, 2);
    if (!name) continue;
    const qty       = Number(v(i, 3)) || 0;
    const unit      = v(i, 4);
    const unitPrice = Number(v(i, 5)) || 0;
    items.push({ name, qty, unit, unitPrice, amount: qty * unitPrice });
  }

  const subtotal = items.reduce((s, it) => s + it.amount, 0);
  const tax      = Math.round(subtotal * 0.1);

  return {
    docType:       v(3, 2),
    issueDate:     v(4, 2),
    dueDate:       sh.getRange(4, 5).getValue(),
    clientName:    v(7, 2),
    clientPerson:  v(8, 2),
    clientZip:     v(9, 2),
    clientAddress: v(10, 2),
    subject:       v(12, 2),
    notes:         v(13, 2),
    items,
    subtotal,
    tax,
    total: subtotal + tax,
  };
}

// ============================================================
// 書類生成メイン
// ============================================================
function generateInvoice() { _generate('請求書'); }
function generateQuote()   { _generate('見積書'); }

function _generate(docType) {
  try {
    const cfg   = _getSettings();
    const input = _getInputData();
    const ui    = SpreadsheetApp.getUi();

    if (!input.clientName) {
      ui.alert('入力エラー', '請求先（会社名/氏名）を入力してください。', ui.ButtonSet.OK);
      return;
    }
    if (input.items.length === 0) {
      ui.alert('入力エラー', '品目を1件以上入力してください。', ui.ButtonSet.OK);
      return;
    }

    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sh   = ss.getSheetByName(SHEET_INVOICE);
    sh.clear(); sh.clearFormats();

    // 列幅設定（A=余白, B=左コンテンツ, C-F=中, G-H=右コンテンツ, I=余白）
    [18, 160, 80, 55, 55, 90, 110, 18].forEach((w, i) => sh.setColumnWidth(i + 1, w));

    const docNo   = _docNumber(cfg.prefix, docType);
    const fmtDate = d => d ? Utilities.formatDate(new Date(d), 'Asia/Tokyo', 'yyyy年MM月dd日') : '';
    const issueStr = input.issueDate ? fmtDate(input.issueDate) : fmtDate(new Date());

    let dueStr = '';
    if (input.dueDate) {
      dueStr = fmtDate(input.dueDate);
    } else {
      const d = new Date(input.issueDate || new Date());
      d.setDate(d.getDate() + cfg.payDays);
      dueStr = fmtDate(d);
    }

    let r = 2; // 開始行

    // ========== タイトル ==========
    _merge(sh, r, 2, 1, 7).setValue(docType)
      .setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground('#1a73e8').setFontColor('white');
    sh.setRowHeight(r, 50);
    r += 2;

    // ========== 書類番号・日付（右寄せ） ==========
    _merge(sh, r, 5, 1, 2).setValue('書　類　番　号').setFontColor('#888888').setHorizontalAlignment('right');
    sh.getRange(r, 7).setValue(docNo).setHorizontalAlignment('right').setFontWeight('bold');
    r++;
    _merge(sh, r, 5, 1, 2).setValue('発　行　日').setFontColor('#888888').setHorizontalAlignment('right');
    sh.getRange(r, 7).setValue(issueStr).setHorizontalAlignment('right');
    r++;

    if (docType === '請求書') {
      _merge(sh, r, 5, 1, 2).setValue('支　払　期　限').setFontColor('#e53935').setHorizontalAlignment('right').setFontWeight('bold');
      sh.getRange(r, 7).setValue(dueStr).setHorizontalAlignment('right').setFontColor('#e53935').setFontWeight('bold');
      r++;
    }
    r++;

    // ========== 請求先 & 発行者 ==========
    const infoTop = r;

    // 左：請求先
    const honorific = input.clientName + (docType === '請求書' ? '　御中' : '　様');
    _merge(sh, r, 2, 1, 4).setValue(honorific)
      .setFontSize(14).setFontWeight('bold')
      .setBorder(false, false, true, false, false, false, '#333333', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    r++;
    if (input.clientZip)     { _merge(sh, r, 2, 1, 4).setValue('〒' + input.clientZip).setFontColor('#555555'); r++; }
    if (input.clientAddress)  { _merge(sh, r, 2, 1, 4).setValue(input.clientAddress).setFontColor('#555555'); r++; }
    if (input.clientPerson)   { _merge(sh, r, 2, 1, 4).setValue(input.clientPerson + '　様').setFontColor('#555555'); r++; }

    // 右：発行者（infoTop から上書き）
    let ir = infoTop;
    const _right = (row, txt, opt = {}) => {
      const cell = _merge(sh, row, 5, 1, 3).setValue(txt).setHorizontalAlignment('right');
      if (opt.bold)  cell.setFontWeight('bold');
      if (opt.size)  cell.setFontSize(opt.size);
      if (opt.color) cell.setFontColor(opt.color);
    };
    if (cfg.companyName) { _right(ir++, cfg.companyName, { bold: true, size: 13 }); }
    if (cfg.ownerName)   { _right(ir++, cfg.ownerName); }
    if (cfg.zip)         { _right(ir++, '〒' + cfg.zip, { color: '#555555' }); }
    if (cfg.address)     { _right(ir++, cfg.address, { color: '#555555' }); }
    if (cfg.phone)       { _right(ir++, 'TEL: ' + cfg.phone, { color: '#555555' }); }
    if (cfg.email)       { _right(ir++, cfg.email, { color: '#555555' }); }
    if (cfg.invoiceReg)  { _right(ir++, '登録番号: ' + cfg.invoiceReg, { color: '#777777' }); }

    r = Math.max(r, ir) + 1;

    // ========== 件名 ==========
    if (input.subject) {
      _merge(sh, r, 2, 1, 6).setValue('件名：' + input.subject)
        .setFontWeight('bold').setFontSize(11).setBackground('#f5f5f5');
      r++;
    }
    r++;

    // ========== 請求金額ボックス ==========
    const amountLabel = docType === '請求書' ? 'ご請求金額' : 'お見積金額';
    _merge(sh, r, 2, 1, 6).setValue(amountLabel)
      .setHorizontalAlignment('center').setFontColor('#666666').setFontSize(10);
    r++;
    _merge(sh, r, 2, 1, 6)
      .setValue('¥ ' + input.total.toLocaleString() + '　（税込）')
      .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground('#e8f0fe')
      .setBorder(true, true, true, true, false, false, '#1a73e8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.setRowHeight(r, 44);
    r += 2;

    // ========== 明細テーブル ==========
    const COL_LABELS = ['品目・内容', '数量', '単位', '単価', '', '金額'];
    // ヘッダー
    sh.getRange(r, 2).setValue(COL_LABELS[0]).setBackground('#1a73e8').setFontColor('white').setFontWeight('bold');
    sh.getRange(r, 3).setValue(COL_LABELS[1]).setBackground('#1a73e8').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(r, 4).setValue(COL_LABELS[2]).setBackground('#1a73e8').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(r, 5).setValue(COL_LABELS[3]).setBackground('#1a73e8').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('right');
    _merge(sh, r, 6, 1, 2).setValue(COL_LABELS[5]).setBackground('#1a73e8').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('right');
    r++;

    input.items.forEach((item, idx) => {
      const bg = idx % 2 === 0 ? '#ffffff' : '#f8f9fa';
      const border = { top: false, left: true, bottom: true, right: true, vertical: false, horizontal: false };
      sh.getRange(r, 2).setValue(item.name).setBackground(bg)
        .setBorder(false, true, true, false, false, false, '#dddddd', SpreadsheetApp.BorderStyle.SOLID);
      sh.getRange(r, 3).setValue(item.qty).setBackground(bg).setHorizontalAlignment('right').setNumberFormat('#,##0')
        .setBorder(false, false, true, false, false, false, '#dddddd', SpreadsheetApp.BorderStyle.SOLID);
      sh.getRange(r, 4).setValue(item.unit).setBackground(bg).setHorizontalAlignment('center')
        .setBorder(false, false, true, false, false, false, '#dddddd', SpreadsheetApp.BorderStyle.SOLID);
      sh.getRange(r, 5).setValue(item.unitPrice).setBackground(bg).setHorizontalAlignment('right').setNumberFormat('¥#,##0')
        .setBorder(false, false, true, false, false, false, '#dddddd', SpreadsheetApp.BorderStyle.SOLID);
      _merge(sh, r, 6, 1, 2).setValue(item.amount).setBackground(bg).setHorizontalAlignment('right').setNumberFormat('¥#,##0')
        .setBorder(false, false, true, true, false, false, '#dddddd', SpreadsheetApp.BorderStyle.SOLID);
      r++;
    });

    // 小計・税・合計
    _merge(sh, r, 2, 1, 5).setValue('小　計').setHorizontalAlignment('right').setFontWeight('bold').setBackground('#f0f4ff');
    sh.getRange(r, 7).setValue(input.subtotal).setHorizontalAlignment('right').setNumberFormat('¥#,##0').setFontWeight('bold').setBackground('#f0f4ff');
    r++;

    _merge(sh, r, 2, 1, 5).setValue('消費税（10%）').setHorizontalAlignment('right').setBackground('#f0f4ff');
    sh.getRange(r, 7).setValue(input.tax).setHorizontalAlignment('right').setNumberFormat('¥#,##0').setBackground('#f0f4ff');
    r++;

    _merge(sh, r, 2, 1, 5).setValue('合計（税込）').setHorizontalAlignment('right').setFontWeight('bold').setFontSize(12)
      .setBackground('#1a73e8').setFontColor('white');
    sh.getRange(r, 7).setValue(input.total).setHorizontalAlignment('right').setNumberFormat('¥#,##0').setFontWeight('bold').setFontSize(12)
      .setBackground('#1a73e8').setFontColor('white');
    r += 2;

    // ========== 振込先（請求書のみ） ==========
    if (docType === '請求書' && cfg.bankName) {
      _merge(sh, r, 2, 1, 6).setValue('【お振込先】').setFontWeight('bold').setBackground('#e8f0fe');
      r++;
      const bankStr = `${cfg.bankName}　${cfg.bankBranch}　${cfg.bankType}　${cfg.bankAccount}　${cfg.bankHolder}`;
      _merge(sh, r, 2, 1, 6).setValue(bankStr).setBackground('#f8f9fa');
      r += 2;
    }

    // ========== 備考 ==========
    if (input.notes) {
      _merge(sh, r, 2, 1, 6).setValue('【備考】').setFontWeight('bold').setBackground('#f5f5f5');
      r++;
      _merge(sh, r, 2, 1, 6).setValue(String(input.notes)).setWrap(true).setBackground('#fafafa');
      sh.setRowHeight(r, 72);
    }

    // 完成→シート移動
    ss.setActiveSheet(sh);
    ui.alert(
      '生成完了！',
      `${docType}（${docNo}）を作成しました。\n\n` +
      'PDF保存：ファイル → ダウンロード → PDF (.pdf)',
      ui.ButtonSet.OK
    );

  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ============================================================
// 書類番号生成
// ============================================================
function _docNumber(prefix, docType) {
  const now   = new Date();
  const ymd   = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  const seq   = String(now.getTime()).slice(-4);
  const type  = docType === '見積書' ? 'Q' : 'I';
  return `${prefix}-${type}${ymd}-${seq}`;
}

// ============================================================
// セル結合ヘルパー（Range を返す）
// ============================================================
function _merge(sh, row, col, numRows, numCols) {
  return sh.getRange(row, col, numRows, numCols).merge();
}

// ============================================================
// 入力シートリセット
// ============================================================
function resetInput() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('確認', '入力内容をすべてクリアしますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INPUT);
  if (!sh) return;

  sh.getRange('B3').setValue('請求書');
  sh.getRange('B4').setValue(new Date());
  sh.getRange('E4').setValue('');

  ['B7','B8','B9','B10','B12','B13'].forEach(a => sh.getRange(a).setValue(''));

  for (let i = 17; i <= 31; i++) {
    sh.getRange(i, 2, 1, 4).setValues([['', '', '', '']]);
  }

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}
