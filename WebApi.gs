/**
 * 品質不具合管理システム Web API
 *
 * 西軽精機アプリ(nishikaru-app、GitHub Pages)からの呼び出し先。
 * 「客先クレーム(CC)・社内不良(KP)管理台帳 2026年度」スプレッドシートのApps Scriptエディタに、
 * SetupSpreadsheet.gs と一緒にこのファイルも追加してデプロイする。
 * 年度が変わると対象スプレッドシートは自動的に切り替わる(getCurrentYearSpreadsheet_、
 * SetupSpreadsheet.gs参照)。GAS自体はどのスプレッドシートに紐づけて作成しても動作に影響しない
 * (SpreadsheetApp.openByIdで都度明示的に開いているため)。
 *
 * 【実行方法】
 * 1. 「客先クレーム(CC)・社内不良(KP)管理台帳 2026年度」スプレッドシートを開く
 * 2. 拡張機能 > Apps Script を開き、新規ファイル「WebApi」を作成してこの内容を貼り付ける
 *    (SetupSpreadsheet.gs は削除せずそのまま残しておく)
 * 3. 【新規追加コード反映時のみ・1回だけ】関数選択で「testExternalFetchAuth」を選び、実行ボタンを押す。
 *    「承認が必要です」という画面が出るので、自分のアカウントを選び「許可」する。
 *    (このコードで初めてUrlFetchAppを使うGoogle以外の外部サイトアクセスの権限を承認する手順。
 *    これをしないとdoPostが「Failed to fetch」で失敗する)
 * 4. 「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. デプロイ後に表示されるURLを、nishikaru-app の index.html 内 GAS_URL 定数に貼り付ける
 * 6. コードを変更するたびに「新しいバージョン」で再デプロイすること(勤怠管理システムと同じ運用)
 *
 * 提供するAPI:
 * - GET  ?mode=lookup&mfgNo=製造番号   → 進捗状況照会シートを検索し、得意先名・品番(図番)を返す(未ログインでも可)
 * - GET  ?mode=masters                 → 組織図マスタから機種名(機械名)・加工者名の一覧を返す(未ログインでも可)
 * - GET  ?mode=dashboard               → ダッシュボード(dashboard.html)用の集計データをJSONで返す(未ログインでも可)
 * - POST { action:'submit', idToken:'...', ... } → 不良〇月シートへ1件書き込む。idTokenをGoogleに
 *   照会して検証できたリクエストのみ受け付け、検証済みのメールアドレスを品証担当者(C列)に記録する。
 */

var PROGRESS_SS_ID = '1F9Iu5t62WDW5lg_eeEa6XW9ngCUJ2DmXTKjqd5oXrac'; // 進捗状況照会(I-Pro連携)
var PROGRESS_SHEET_NAME = '進捗状況照会';
var ORG_MASTER_SS_ID = '1fffjE_bwrzswvRO62U0OHwvqrs5b_UuSV5IbudUMxec'; // 組織図マスタ
var ORG_MASTER_SHEET_NAME = 'プルダウン用';

// Google Identity Services(ログイン)用。品質保証課アプリ用に発行したクライアントID
var OAUTH_CLIENT_ID = '800178947678-qdka8ic7v2c5bbeocgiqd8qqrheafq7e.apps.googleusercontent.com';
var ALLOWED_EMAIL_DOMAIN = 'nishikaru.co.jp';

var DATA_START_ROW = 2;
var DATA_END_ROW = 115;

function doGet(e) {
  var mode = e.parameter.mode;
  if (mode === 'lookup') {
    return jsonOutput_(lookupByMfgNo_(e.parameter.mfgNo || ''));
  }
  if (mode === 'masters') {
    return jsonOutput_(getMasters_());
  }
  if (mode === 'dashboard') {
    return jsonOutput_(buildDashboardData_());
  }
  return jsonOutput_({ error: 'unknown mode: ' + mode });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ ok: false, error: 'リクエスト内容を解析できませんでした' });
  }

  var action = body.action || 'submit';
  if (action === 'submit') {
    var verified = verifyGoogleToken_(body.idToken);
    if (!verified) {
      return jsonOutput_({ ok: false, error: 'ログインを確認できませんでした。再読み込みしてログインし直してください。' });
    }
    try {
      writeDefectRecord_(body, verified.name || verified.email);
      return jsonOutput_({ ok: true });
    } catch (err) {
      return jsonOutput_({ ok: false, error: err.message });
    }
  }
  return jsonOutput_({ ok: false, error: 'unknown action: ' + action });
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 【1回だけ手動実行】UrlFetchAppでGoogle以外の外部サイト(oauth2.googleapis.com)へアクセスする
 * 権限を承認させるためだけの関数。GASエディタでこの関数を選んで実行ボタンを押すと
 * 「承認が必要です」ダイアログが出るので許可する。承認済みなら何度実行しても無害。
 * (名前の末尾にアンダースコアを付けない: GASの「実行」プルダウンは末尾が_の関数を表示しないため、
 * 手動実行してほしいこの関数だけは意図的にアンダースコアなしにしてある)
 */
function testExternalFetchAuth() {
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=test', { muteHttpExceptions: true });
  Logger.log('外部サイトアクセスの権限は問題ありません(応答コード: ' + res.getResponseCode() + ')');
}

/**
 * クライアントから渡されたGoogle IDトークンをGoogleのtokeninfoエンドポイントに照会して検証する。
 * 検証OKなら { email, name } を返し、NGならnullを返す(クライアント側の申告値は信用しない)。
 * 勤怠申請アプリ(shuusei-app)のverifyGoogleToken_と同じ考え方。
 */
function verifyGoogleToken_(idToken) {
  if (!idToken) return null;
  var res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) return null;

  var payload;
  try {
    payload = JSON.parse(res.getContentText());
  } catch (err) {
    return null;
  }

  if (payload.aud !== OAUTH_CLIENT_ID) return null;
  if (!payload.email || payload.email_verified !== 'true') return null;
  if (ALLOWED_EMAIL_DOMAIN && payload.email.split('@')[1] !== ALLOWED_EMAIL_DOMAIN) return null;

  return { email: payload.email, name: payload.name || '' };
}

/**
 * 製造番号から「進捗状況照会」シートを検索し、得意先名・品番(図番)を返す。
 * 同じ製造番号は複数行(工程ごと)にまたがることがあるが、得意先名・品番はどの行でも同じなので
 * 最初に見つかった行を採用する。
 */
function lookupByMfgNo_(mfgNo) {
  mfgNo = (mfgNo || '').toString().trim();
  if (!mfgNo) return { found: false, error: '製造番号が空です' };

  var ss = SpreadsheetApp.openById(PROGRESS_SS_ID);
  var sheet = ss.getSheetByName(PROGRESS_SHEET_NAME);
  if (!sheet) return { found: false, error: '「' + PROGRESS_SHEET_NAME + '」シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var colMfgNo = header.indexOf('製造番号');
  var colCustomer = header.indexOf('得意先名');
  var colDrawing = header.indexOf('品番(図番)');
  if (colMfgNo === -1 || colCustomer === -1 || colDrawing === -1) {
    return { found: false, error: '想定した列(製造番号/得意先名/品番(図番))が見つかりません' };
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][colMfgNo].toString().trim() === mfgNo) {
      return {
        found: true,
        mfgNo: mfgNo,
        customer: data[i][colCustomer].toString(),
        drawing: data[i][colDrawing].toString()
      };
    }
  }
  return { found: false, mfgNo: mfgNo };
}

/**
 * 「組織図マスタ」の「プルダウン用」シートから機種名(機械名列)・加工者名(加工者名列)の
 * 一覧を取り出す(重複を除いて登場順)。西軽精機アプリの入力フォームのプルダウンに使う。
 */
function getMasters_() {
  var ss = SpreadsheetApp.openById(ORG_MASTER_SS_ID);
  var sheet = ss.getSheetByName(ORG_MASTER_SHEET_NAME);
  if (!sheet) return { kishu: [], kakosha: [], error: '「' + ORG_MASTER_SHEET_NAME + '」シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var colKishu = header.indexOf('機械名');
  var colKakosha = header.indexOf('加工者名');
  if (colKishu === -1 || colKakosha === -1) {
    return { kishu: [], kakosha: [], error: '想定した列(機械名/加工者名)が見つかりません' };
  }

  var kishu = [];
  var kakosha = [];
  for (var i = 1; i < data.length; i++) {
    var k = data[i][colKishu].toString().trim();
    var p = data[i][colKakosha].toString().trim();
    if (k && kishu.indexOf(k) === -1) kishu.push(k);
    if (p && kakosha.indexOf(p) === -1) kakosha.push(p);
  }
  return { kishu: kishu, kakosha: kakosha };
}

/**
 * 不良〇月シートへ1件書き込む。
 * 列構成(SetupSpreadsheet.gsの buildDefectMonthlySheet_ と対応させること):
 * A タイムスタンプ／B 処置区分／C 品証担当者／D 製造番号／E 得意先名／F 品番(図番)／G 加工者／
 * H 機種名／I 設備№／J 加工数／K 良品数(自動計算式、書き込まない)／L 不良数計／M 不良項目／
 * N 不良数／O 不良項目詳細／P 担当者2(不良項目ごとに原因を作った担当者が違う場合のみ入力する任意項目)／
 * Q 単価／R 金額(自動計算式、書き込まない)／S 備考／T 材質／U 不良率(自動計算式、書き込まない)／
 * V キズ原因(任意)／W 送信ID(今回未使用)
 * 【2026-08-13改訂】D列「製造番号」を新設(QRスキャンで取得済みだったがシートに保存していなかった)。
 * これによりD列以降が1列ずつ後ろにずれている。
 *
 * K・R・U列はSetupSpreadsheet.gs側で数式を全行にあらかじめ設定してあるため、
 * ここで値を書き込むと数式が消えてしまう。書き込み対象からは常に除外する。
 *
 * 不良項目が複数ある場合、1件目はメイン行のM〜P列に、2件目以降はB列(処置区分。行の色分けを
 * 効かせるため)・M〜P列(不良項目・不良数・詳細・担当者2)だけの追加行に書く。
 * 書き込み終わったら、その1件分の行(複数行にまたがる場合はまとめて)を枠線で囲み、次の送信との
 * 区切りが分かりやすいようにする(旧システムのtransferToMonthlySheetと同じ考え方)。
 *
 * @param {Object} body リクエストの中身。items は [{name, qty, detail, worker2}] の配列。
 * @param {string} verifiedName verifyGoogleToken_で確認済みの氏名(取得できなければメール)。
 *   C列(品証担当者)にそのまま使う(クライアントが送ってきた値ではなく、サーバー側で検証済みの値を信用する)。
 */
function writeDefectRecord_(body, verifiedName) {
  var now = new Date();
  var ss = getCurrentYearSpreadsheet_(now); // 年度自動ロールオーバー対応(SetupSpreadsheet.gs参照、同一GASプロジェクト内で共有)
  var month = now.getMonth() + 1;
  var sheetName = '不良' + month + '月';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('シート「' + sheetName + '」が見つかりません');

  var items = Array.isArray(body.items) ? body.items : [];
  if (items.length === 0) throw new Error('不良項目が指定されていません');

  var totalQty = items.reduce(function (sum, it) { return sum + (Number(it.qty) || 0); }, 0);

  var row = findNextRow_(sheet);
  if (row > DATA_END_ROW) throw new Error(sheetName + ' が満杯です');

  sheet.getRange(row, 1, 1, 10).setValues([[
    now,                          // A タイムスタンプ
    body.shochiKubun || '',       // B 処置区分
    verifiedName,                  // C 品証担当者(サーバーで検証済みの氏名)
    body.mfgNo || '',             // D 製造番号
    body.customer || '',          // E 得意先名
    body.drawing || '',           // F 品番(図番)
    body.kakosha || '',           // G 加工者
    body.kishu || '',             // H 機種名
    body.setsubi || '',           // I 設備№
    Number(body.suryo) || ''      // J 加工数
  ]]);
  sheet.getRange(row, 12, 1, 5).setValues([[
    totalQty,                     // L 不良数計
    items[0].name,                // M 不良項目(1件目)
    Number(items[0].qty) || '',   // N 不良数(1件目)
    items[0].detail || '',        // O 不良項目詳細(1件目)
    items[0].worker2 || ''        // P 担当者2(1件目、任意)
  ]]);
  sheet.getRange(row, 17).setValue(Number(body.tanka) || '');  // Q 単価
  sheet.getRange(row, 19).setValue(body.biko || '');           // S 備考
  sheet.getRange(row, 20).setValue(body.zaishitsu || '');      // T 材質
  sheet.getRange(row, 22).setValue(body.kizugenin || '');      // V キズ原因(任意)

  for (var i = 1; i < items.length; i++) {
    var r = row + i;
    if (r > DATA_END_ROW) throw new Error(sheetName + ' が満杯です(不良項目の追加行)');
    sheet.getRange(r, 2).setValue(body.shochiKubun || ''); // B 処置区分(追加行にも複製、行の色分け用)
    sheet.getRange(r, 13, 1, 4).setValues([[
      items[i].name, Number(items[i].qty) || '', items[i].detail || '', items[i].worker2 || ''
    ]]);
  }

  // 1件の入力(不良項目が複数で複数行にまたがる場合はまとめて)を枠線で囲み、次の入力と見分けやすくする。
  // 旧システム(検査不具合報告\コード.js)のtransferToMonthlySheetが送信のたびに行っていたのと同じ考え方。
  // 外枠のみ(内部に縦線・横線は引かない)。
  sheet.getRange(row, 1, items.length, 23)
    .setBorder(true, true, true, true, false, false, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * ダッシュボード(dashboard.html)用の集計データを組み立てる。
 * MONTHS・DEFECT_ITEMS・KP_CAUSE_ITEMS・uniqueInOrder_ は同じGASプロジェクトの
 * SetupSpreadsheet.gs で定義済みのものをそのまま使う(同一プロジェクト内はグローバル共有のため、
 * writeDefectRecord_ が COLOR を参照しているのと同じ考え方)。
 * 「不良集計」「不良集計(キズ原因)」「月次サマリー」「クレーム集計」の列構成は SetupSpreadsheet.gs の
 * buildItemSummarySheet_ / buildMonthlySummarySheet_ / buildClaimSummarySheet_ と対応させること
 * (ずれると集計が崩れる)。
 */
function buildDashboardData_() {
  var ss = getCurrentYearSpreadsheet_(); // 年度自動ロールオーバー対応(SetupSpreadsheet.gs参照、同一GASプロジェクト内で共有)

  // --- 月次サマリー: 月別 個数・件数・金額(KP)、KP/差し戻しの年計個数 ---
  var summarySheet = ss.getSheetByName('月次サマリー');
  var monthlyQty = summarySheet.getRange(10, 2, 1, MONTHS.length).getValues()[0].map(Number);   // 合計 不良個数(KP+差し戻し)
  var monthlyCount = summarySheet.getRange(9, 2, 1, MONTHS.length).getValues()[0].map(Number);  // 合計 不良件数(KP+差し戻し)
  var monthlyAmountKP = summarySheet.getRange(4, 2, 1, MONTHS.length).getValues()[0].map(Number); // KP 不良金額
  var kpQtyYear = Number(summarySheet.getRange(3, 14).getValue()) || 0;     // KP 不良個数(年計)
  var reworkQtyYear = Number(summarySheet.getRange(7, 14).getValue()) || 0; // 差し戻し 個数(年計)

  // --- 不良集計: 分類別×月別の個数(SUMIF相当をJS側で計算) ---
  var defectGroups = uniqueInOrder_(DEFECT_ITEMS.map(function (item) { return item.group; }));
  var itemSheet = ss.getSheetByName('不良集計');
  var itemGroups = itemSheet.getRange(3, 2, DEFECT_ITEMS.length, 1).getValues().map(function (r) { return r[0]; });
  var stackedByGroup = MONTHS.map(function (month, mi) {
    var qtyCol = 4 + mi * 2; // 不良集計シートの月別「個数」列(D,F,H...)
    var qtyValues = itemSheet.getRange(3, qtyCol, DEFECT_ITEMS.length, 1).getValues().map(function (r) { return Number(r[0]) || 0; });
    return defectGroups.map(function (g) {
      var sum = 0;
      for (var i = 0; i < itemGroups.length; i++) if (itemGroups[i] === g) sum += qtyValues[i];
      return sum;
    });
  });

  // --- 不良集計(キズ原因): 原因グループ別の年計個数 ---
  var causeGroups = uniqueInOrder_(KP_CAUSE_ITEMS.map(function (item) { return item.group; }));
  var causeSheet = ss.getSheetByName('不良集計(キズ原因)');
  var causeSheetGroups = causeSheet.getRange(3, 2, KP_CAUSE_ITEMS.length, 1).getValues().map(function (r) { return r[0]; });
  var causeYearQtyCol = 4 + MONTHS.length * 2 + 1; // 「年計」個数列
  var causeQtyValues = causeSheet.getRange(3, causeYearQtyCol, KP_CAUSE_ITEMS.length, 1).getValues().map(function (r) { return Number(r[0]) || 0; });
  var causeTotals = causeGroups.map(function (g) {
    var sum = 0;
    for (var i = 0; i < causeSheetGroups.length; i++) if (causeSheetGroups[i] === g) sum += causeQtyValues[i];
    return sum;
  });

  // --- クレーム集計: 客先クレーム件数の月別合計(合計行、SetupSpreadsheet.gsのbuildClaimSummarySheet_と対応) ---
  var claimSheet = ss.getSheetByName('クレーム集計');
  var claimMonthly = claimSheet
    ? claimSheet.getRange(2 + DEFECT_ITEMS.length, 3, 1, MONTHS.length).getValues()[0].map(Number)
    : MONTHS.map(function () { return 0; });

  // --- 客先クレーム管理台帳(CC): 加工者別・検査員別のクレーム件数(2026-08-17追加) ---
  var kakoshaCount = {}, kensainCount = {};
  var ccSheetForRanking = ss.getSheetByName(CC_LEDGER_SHEET_NAME);
  if (ccSheetForRanking) {
    var ccLastRow = ccSheetForRanking.getLastRow();
    var ccRows = Math.max(ccLastRow - CC_DATA_START_ROW + 1, 0);
    if (ccRows > 0) {
      var kakoshaValues = ccSheetForRanking.getRange(CC_DATA_START_ROW, CC_WORKER_COL, ccRows, 1).getValues();
      var kensainValues = ccSheetForRanking.getRange(CC_DATA_START_ROW, CC_INSPECTOR_COL, ccRows, 1).getValues();
      kakoshaValues.forEach(function (r) {
        var v = r[0] ? r[0].toString().trim() : '';
        if (v) kakoshaCount[v] = (kakoshaCount[v] || 0) + 1;
      });
      kensainValues.forEach(function (r) {
        var v = r[0] ? r[0].toString().trim() : '';
        if (v) kensainCount[v] = (kensainCount[v] || 0) + 1;
      });
    }
  }

  // --- 不良〇月シート12枚: 得意先別金額・設備別個数・月別加工数合計(不良率の分母) ---
  var customerAmount = {}; // { 得意先名: 金額合計 }
  var machineQty = {};     // { 機種名: 不良個数合計 }
  var monthlyVolume = [];  // 月別 加工数合計

  MONTHS.forEach(function (month) {
    var sheet = ss.getSheetByName('不良' + month + '月');
    var rows = sheet.getRange(2, 1, 114, 18).getValues(); // A〜R列(E得意先名/H機種名/J加工数/L不良数計/R金額)
    var volume = 0;
    rows.forEach(function (row) {
      var customer = row[4], machine = row[7], suryo = row[9], qtyTotal = row[11], amount = row[17];
      if (suryo) volume += Number(suryo) || 0;
      if (machine && qtyTotal) machineQty[machine] = (machineQty[machine] || 0) + (Number(qtyTotal) || 0);
      if (customer && amount) customerAmount[customer] = (customerAmount[customer] || 0) + (Number(amount) || 0);
    });
    monthlyVolume.push(volume);
  });

  var defectRate = MONTHS.map(function (month, mi) {
    var v = monthlyVolume[mi];
    return v > 0 ? Math.round((monthlyQty[mi] / v) * 1000) / 10 : 0;
  });

  var kpReworkTotal = kpQtyYear + reworkQtyYear;
  var kpReworkRatio = kpReworkTotal > 0 ? [
    { label: '社内不良(KP)', value: Math.round((kpQtyYear / kpReworkTotal) * 1000) / 10 },
    { label: '差し戻し', value: Math.round((reworkQtyYear / kpReworkTotal) * 1000) / 10 }
  ] : [];

  function topN(obj, n) {
    return Object.keys(obj)
      .map(function (k) { return { label: k, value: obj[k] }; })
      .sort(function (a, b) { return b.value - a.value; })
      .slice(0, n);
  }

  return {
    months: MONTHS.map(function (m) { return m + '月'; }),
    monthlyQty: monthlyQty,
    monthlyCount: monthlyCount,
    monthlyAmountKP: monthlyAmountKP,
    defectGroups: defectGroups,
    stackedByGroup: stackedByGroup,
    causeGroups: causeGroups,
    causeTotals: causeTotals,
    customers: topN(customerAmount, 8),
    machines: topN(machineQty, 8),
    defectRate: defectRate,
    kpReworkRatio: kpReworkRatio,
    claimMonthly: claimMonthly,
    claimByWorker: topN(kakoshaCount, 8),
    claimByInspector: topN(kensainCount, 8),
    updatedAt: new Date().toISOString()
  };
}

/** シートの次の空き行を探す(A列・L列のどちらかが埋まっていればその行は使用済みとみなす) */
function findNextRow_(sheet) {
  var rowCount = DATA_END_ROW - DATA_START_ROW + 1;
  var colA = sheet.getRange(DATA_START_ROW, 1, rowCount, 1).getValues();
  var colM = sheet.getRange(DATA_START_ROW, 13, rowCount, 1).getValues(); // M列(不良項目)
  var last = DATA_START_ROW - 1;
  for (var i = 0; i < colA.length; i++) {
    if (colA[i][0] !== '' || colM[i][0] !== '') last = DATA_START_ROW + i;
  }
  return last + 1;
}
