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
 * - GET  ?mode=lookup&mfgNo=製造番号   → 進捗状況照会シートを検索し、得意先名・品番(図番)・機種名・
 *   設備№・加工者名・加工数・材質・単価を返す(未ログインでも可、2026-08-19に得意先名・品番のみから拡張)
 * - GET  ?mode=masters                 → 組織図マスタから機種名(機械名)・加工者名の一覧を返す(未ログインでも可)
 * - GET  ?mode=dashboard               → ダッシュボード(dashboard.html)用の集計データをJSONで返す(未ログインでも可)
 * - POST { action:'submit', idToken:'...', ... } → 不良〇月シートへ1件書き込む。idTokenをGoogleに
 *   照会して検証できたリクエストのみ受け付け、検証済みのメールアドレスを品証担当者(C列)に記録する。
 * - POST { action:'lookupRecord', idToken:'...', mfgNo:'...' } → その製造番号について、ログイン中の
 *   本人が過去に送信した記録(今月・先月の「不良〇月」シートのみ検索)を探し、見つかれば編集用に
 *   全項目を返す(2026-08-18新設、編集機能)。
 * - POST { action:'update', idToken:'...', id:'...', month:数値, ... } → lookupRecordで見つけた記録を
 *   同じ送信ID(W列)の行に上書きする(2026-08-18新設)。本人の記録以外は更新を拒否する。
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
      var newId = writeDefectRecord_(body, verified.name || verified.email);
      return jsonOutput_({ ok: true, id: newId });
    } catch (err) {
      return jsonOutput_({ ok: false, error: err.message });
    }
  }

  if (action === 'lookupRecord') {
    var verifiedForLookup = verifyGoogleToken_(body.idToken);
    if (!verifiedForLookup) {
      return jsonOutput_({ found: false, error: 'ログインを確認できませんでした。再読み込みしてログインし直してください。' });
    }
    try {
      return jsonOutput_(lookupRecordForEdit_(body.mfgNo || '', verifiedForLookup.name || verifiedForLookup.email));
    } catch (err) {
      return jsonOutput_({ found: false, error: err.message });
    }
  }

  if (action === 'update') {
    var verifiedForUpdate = verifyGoogleToken_(body.idToken);
    if (!verifiedForUpdate) {
      return jsonOutput_({ ok: false, error: 'ログインを確認できませんでした。再読み込みしてログインし直してください。' });
    }
    try {
      updateDefectRecord_(body, verifiedForUpdate.name || verifiedForUpdate.email);
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
 * 製造番号から「進捗状況照会」シートを検索し、フォーム自動入力用の項目を返す(2026-08-19拡張)。
 * 同じ製造番号は複数行(工程ごと)にまたがるため、項目によって採用する行を使い分ける:
 * - 得意先名・品番(図番)・材質(品名・型格)・単価(納入単価): どの行でも同じ値のはずなので、
 *   最初に見つかった行を採用する。
 * - 機種名(設備名)・設備№(設備No.)・加工者名(担当者名): 工程順が1の行を採用する
 *   (以降の工程で設備・担当者が変わることもあるため、あくまで入力の初期値。アプリ側では
 *   引き続き変更可能な入力欄のまま)。
 * - 加工数(完了数量): 値が入っている工程順の中で一番大きい(最後に記録された)行を採用する。
 * 材質は「ハイフン/×より後ろ」「カッコとその中身」「余分なスペース」を取り除いて返す
 * (`cleanMaterialName_`参照、品名・型格の生値は「SUS303-G」のように末尾に加工方法等の
 * 付随情報が付いていることがあり、材質そのものだけを欲しいため)。
 */
function lookupByMfgNo_(mfgNo) {
  mfgNo = (mfgNo || '').toString().trim();
  if (!mfgNo) return { found: false, error: '製造番号が空です' };

  var ss = SpreadsheetApp.openById(PROGRESS_SS_ID);
  var sheet = ss.getSheetByName(PROGRESS_SHEET_NAME);
  if (!sheet) return { found: false, error: '「' + PROGRESS_SHEET_NAME + '」シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var col = {
    mfgNo: header.indexOf('製造番号'),
    customer: header.indexOf('得意先名'),
    drawing: header.indexOf('品番(図番)'),
    processOrder: header.indexOf('工程順'),
    workerName: header.indexOf('担当者名'),
    completedQty: header.indexOf('完了数量'),
    equipmentName: header.indexOf('設備名'),
    equipmentNo: header.indexOf('設備No.'),
    material: header.indexOf('品名・型格'),
    unitPrice: header.indexOf('納入単価')
  };
  if (col.mfgNo === -1 || col.customer === -1 || col.drawing === -1) {
    return { found: false, error: '想定した列(製造番号/得意先名/品番(図番))が見つかりません' };
  }

  var result = null;
  var firstProcessRow = null; // 工程順=1の行(機種名・設備№・加工者名の候補)
  var lastQtyRow = null;      // 完了数量が入っている中で工程順が一番大きい行(加工数の候補)

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[col.mfgNo].toString().trim() !== mfgNo) continue;

    if (!result) {
      result = {
        found: true,
        mfgNo: mfgNo,
        customer: row[col.customer].toString(),
        drawing: row[col.drawing].toString(),
        zaishitsu: col.material !== -1 ? cleanMaterialName_(row[col.material]) : '',
        tanka: col.unitPrice !== -1 ? row[col.unitPrice] : ''
      };
    }

    var processOrder = col.processOrder !== -1 ? Number(row[col.processOrder]) : null;
    if (processOrder === 1 && !firstProcessRow) firstProcessRow = row;

    var completedQty = col.completedQty !== -1 ? row[col.completedQty] : '';
    if (completedQty !== '' && completedQty !== null && processOrder !== null) {
      if (!lastQtyRow || processOrder > Number(lastQtyRow[col.processOrder])) lastQtyRow = row;
    }
  }

  if (!result) return { found: false, mfgNo: mfgNo };

  if (firstProcessRow) {
    result.kishu = col.equipmentName !== -1 ? firstProcessRow[col.equipmentName].toString() : '';
    result.setsubi = col.equipmentNo !== -1 ? firstProcessRow[col.equipmentNo].toString() : '';
    result.kakosha = col.workerName !== -1 ? firstProcessRow[col.workerName].toString() : '';
  }
  if (lastQtyRow) {
    result.suryo = col.completedQty !== -1 ? lastQtyRow[col.completedQty] : '';
  }

  return result;
}

/**
 * 「品名・型格」の生値から材質名だけを取り出す(2026-08-19新設)。
 * 例: "SUS303-G" → "SUS303"、"SUS303(黒処理)" → "SUS303"、"SUS303 - G " → "SUS303"
 * ハイフン(-)・×より後ろは不要、カッコ(全角・半角とも)とその中身も不要、
 * 前後・内部の余分なスペースは削除して詰める、という指定。
 */
function cleanMaterialName_(raw) {
  var s = (raw || '').toString();
  s = s.replace(/（[^）]*）/g, '').replace(/\([^)]*\)/g, ''); // カッコと中身を除去(先に処理し、カッコ内のハイフンに惑わされないようにする)
  var cutIndex = s.length;
  var hyphenIdx = s.indexOf('-');
  if (hyphenIdx !== -1) cutIndex = Math.min(cutIndex, hyphenIdx);
  var xIdx = s.indexOf('×');
  if (xIdx !== -1) cutIndex = Math.min(cutIndex, xIdx);
  s = s.slice(0, cutIndex);
  s = s.replace(/[\s　]/g, ''); // 前後・内部の余分なスペースを削除して詰める
  return s;
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
 * H 機種名／I 設備№／J 加工数／K 良品数(自動計算式、書き込まない)／L 修正数(記録全体で1つ、
 * 差し戻し時のみ入力)／M 不良数計／N 不良項目／O キズ原因／P 不良数／Q 不良項目詳細／
 * R 担当者2(不良項目ごとに原因を作った担当者が違う場合のみ入力する任意項目)／S 単価／
 * T 金額(自動計算式、書き込まない)／U 備考／V 材質／W 不良率(自動計算式、書き込まない)／X 送信ID
 * 【2026-08-13改訂】D列「製造番号」を新設(QRスキャンで取得済みだったがシートに保存していなかった)。
 * これによりD列以降が1列ずつ後ろにずれている。
 * 【2026-08-18改訂】キズ原因をV列からN列(不良項目の直後)へ移動し、常時表示に変更(以前は非表示・
 * 任意項目の位置づけだったが、入力アプリ側で条件表示・場合により必須の実項目になったため)。
 * これによりN列以降(旧「不良数」以降)が1列ずつ後ろにずれている。
 * 【2026-08-19改訂】K列(良品数)の右にL列「修正数」を新設(差し戻し入力時、いくつ修正できたかを
 * 記録全体で1つ記録する項目。良品数の自動計算式には使わない単純な記録項目)。これによりL列以降
 * (旧「不良数計」以降)が1列ずつ後ろにずれている。あわせて、不良項目ごとの数量欄(旧O列、現P列)の
 * アプリ側プレースホルダを「数量」から「不良数」に統一(処置区分によらず同じ表記にする)。
 *
 * K・T・W列はSetupSpreadsheet.gs側で数式を全行にあらかじめ設定してあるため、
 * ここで値を書き込むと数式が消えてしまう。書き込み対象からは常に除外する。
 *
 * 不良項目が複数ある場合、1件目はメイン行のM・O〜Q列に、2件目以降はB列(処置区分。行の色分けを
 * 効かせるため)・M・O〜Q列(不良項目・不良数・詳細・担当者2)だけの追加行に書く。キズ原因(N列)は
 * 1件の送信につき1つの値(レコード全体で共通)のため、メイン行にだけ書く。
 * 書き込み終わったら、その1件分の行(複数行にまたがる場合はまとめて)を枠線で囲み、次の送信との
 * 区切りが分かりやすいようにする(旧システムのtransferToMonthlySheetと同じ考え方)。
 *
 * @param {Object} body リクエストの中身。items は [{name, qty, detail, worker2}] の配列。
 * @param {string} verifiedName verifyGoogleToken_で確認済みの氏名(取得できなければメール)。
 *   C列(品証担当者)にそのまま使う(クライアントが送ってきた値ではなく、サーバー側で検証済みの値を信用する)。
 * @return {string} 書き込んだ記録の送信ID(W列、後から編集(action=update)で見つけるためのキー)
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

  var row = findNextRow_(sheet);
  if (row + items.length - 1 > DATA_END_ROW) throw new Error(sheetName + ' が満杯です');

  var submissionId = Utilities.getUuid();
  writeRecordRows_(sheet, row, body, verifiedName, submissionId, items, now);
  return submissionId;
}

/**
 * 指定した開始行から、1件分のレコード(不良項目が複数なら複数行)を書き込む共通処理
 * (2026-08-18、writeDefectRecord_とupdateDefectRecord_で共有するために切り出した)。
 * K・T・W列(良品数・金額・不良率)はSetupSpreadsheet.gs側の数式のため書き込み対象から常に除外する。
 * @param {Date} timestamp A列(タイムスタンプ)に書く値。新規送信は現在時刻、更新は元の送信時刻を維持する。
 */
function writeRecordRows_(sheet, startRow, body, verifiedName, submissionId, items, timestamp) {
  var totalQty = items.reduce(function (sum, it) { return sum + (Number(it.qty) || 0); }, 0);

  sheet.getRange(startRow, 1, 1, 10).setValues([[
    timestamp,                     // A タイムスタンプ
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
  sheet.getRange(startRow, 12).setValue(Number(body.shuseisu) || ''); // L 修正数(記録全体で1つ、差し戻し時のみ意味を持つ)
  sheet.getRange(startRow, 13, 1, 2).setValues([[
    totalQty,                     // M 不良数計
    items[0].name                 // N 不良項目(1件目)
  ]]);
  sheet.getRange(startRow, 15).setValue(body.kizugenin || '');      // O キズ原因(レコード全体で1つ)
  sheet.getRange(startRow, 16, 1, 3).setValues([[
    Number(items[0].qty) || 0,    // P 不良数(1件目)。0も有効な値のため空欄にはしない(2026-08-19)
    items[0].detail || '',        // Q 不良項目詳細(1件目)
    items[0].worker2 || ''        // R 担当者2(1件目、任意)
  ]]);
  sheet.getRange(startRow, 19).setValue(Number(body.tanka) || '');  // S 単価
  sheet.getRange(startRow, 21).setValue(body.biko || '');           // U 備考
  sheet.getRange(startRow, 22).setValue(body.zaishitsu || '');      // V 材質
  sheet.getRange(startRow, 24).setValue(submissionId);              // X 送信ID

  for (var i = 1; i < items.length; i++) {
    var r = startRow + i;
    sheet.getRange(r, 2).setValue(body.shochiKubun || ''); // B 処置区分(追加行にも複製、行の色分け用)
    sheet.getRange(r, 14).setValue(items[i].name);         // N 不良項目(追加行)
    sheet.getRange(r, 16, 1, 3).setValues([[
      Number(items[i].qty) || 0, items[i].detail || '', items[i].worker2 || ''
    ]]); // P・Q・R(不良数・詳細・担当者2、追加行)。不良数は0も有効な値のため空欄にはしない
    sheet.getRange(r, 24).setValue(submissionId); // X 送信ID(追加行にも複製、編集時にまとめて見つけるため)
  }

  // 1件の入力(不良項目が複数で複数行にまたがる場合はまとめて)を枠線で囲み、次の入力と見分けやすくする。
  // 旧システム(検査不具合報告\コード.js)のtransferToMonthlySheetが送信のたびに行っていたのと同じ考え方。
  // 外枠のみ(内部に縦線・横線は引かない)。
  sheet.getRange(startRow, 1, items.length, 24)
    .setBorder(true, true, true, true, false, false, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * 製造番号+ログイン中の本人から、過去に送信した記録(編集対象)を探す(2026-08-18新設)。
 * 今月・先月の「不良〇月」シートだけを検索する(月をまたいで全12シートを毎回走査すると重いため。
 * 通常、直近の入力ミスを直したいケースを想定しており、古い月の記録は対象外でよいとの前提)。
 * 同じ製造番号・本人の記録が複数あれば、タイムスタンプが一番新しいものを返す。
 * @return {{found:boolean, id?:string, month?:number, record?:Object}}
 */
function lookupRecordForEdit_(mfgNo, verifiedName) {
  mfgNo = (mfgNo || '').toString().trim();
  if (!mfgNo) return { found: false };

  var now = new Date();
  var ss = getCurrentYearSpreadsheet_(now);
  var thisMonth = now.getMonth() + 1;
  var prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var prevMonth = prevMonthDate.getMonth() + 1;

  var best = null; // { month, row, timestamp }
  [thisMonth, prevMonth].forEach(function (month) {
    var sheet = ss.getSheetByName('不良' + month + '月');
    if (!sheet) return;
    var rows = DATA_END_ROW - DATA_START_ROW + 1;
    var values = sheet.getRange(DATA_START_ROW, 1, rows, 24).getValues(); // A〜X列
    for (var i = 0; i < values.length; i++) {
      var v = values[i];
      var timestamp = v[0]; // A タイムスタンプ(メイン行のみ入っている)
      var mfgNoCell = v[3]; // D 製造番号
      var quality = v[2];   // C 品証担当者
      if (!timestamp || mfgNoCell !== mfgNo || quality !== verifiedName) continue;
      if (!best || timestamp > best.timestamp) {
        best = { month: month, row: DATA_START_ROW + i, timestamp: timestamp };
      }
    }
  });
  if (!best) return { found: false };

  var sheet = ss.getSheetByName('不良' + best.month + '月');
  var id = sheet.getRange(best.row, 24).getValue();
  if (!id) return { found: false }; // この機能が入る前に送信された記録(送信IDが無い)は編集対象外

  var rows = findRecordRows_(sheet, id);
  return { found: true, id: id, month: best.month, record: readRecordFromRows_(sheet, rows) };
}

/** シートのW列(送信ID)から、指定IDに一致する行番号をすべて返す(昇順) */
function findRecordRows_(sheet, id) {
  var rows = DATA_END_ROW - DATA_START_ROW + 1;
  var idValues = sheet.getRange(DATA_START_ROW, 24, rows, 1).getValues();
  var matched = [];
  for (var i = 0; i < idValues.length; i++) {
    if (idValues[i][0] === id) matched.push(DATA_START_ROW + i);
  }
  return matched;
}

/** 行番号の配列(1件分、メイン行+不良項目の追加行)から、フォーム編集用のレコードを再構築する */
function readRecordFromRows_(sheet, rows) {
  if (rows.length === 0) return null;
  var mainRow = rows[0];
  var main = sheet.getRange(mainRow, 1, 1, 24).getValues()[0];

  var items = rows.map(function (row) {
    var m = (row === mainRow) ? main : sheet.getRange(row, 1, 1, 24).getValues()[0];
    return { name: m[13], qty: m[15], detail: m[16] || '', worker2: m[17] || '' }; // N・P・Q・R列
  }).filter(function (it) { return it.name; });

  return {
    shochiKubun: main[1],   // B
    mfgNo: main[3],         // D
    customer: main[4],      // E
    drawing: main[5],       // F
    kakosha: main[6],       // G
    kishu: main[7],         // H
    setsubi: main[8],       // I
    suryo: main[9],         // J
    shuseisu: main[11],     // L 修正数
    kizugenin: main[14],    // O
    tanka: main[18],        // S
    biko: main[20],         // U
    zaishitsu: main[21],    // V
    items: items
  };
}

/**
 * lookupRecordForEdit_で見つけた記録を上書きする(2026-08-18新設)。
 * 本人の記録以外は更新できないようサーバー側でも品証担当者を照合する(クライアント側の制御だけに頼らない)。
 * 送信ID(W列)はそのまま維持する(タイムスタンプも新規送信扱いにはせず、元の値を保つ)。
 * 新しい不良項目数が元の行数より多い場合は、直後の行が本当に空いているか確認してから使う
 * (別の記録の行を誤って上書きしないための安全策)。
 */
function updateDefectRecord_(body, verifiedName) {
  if (!body.id) throw new Error('編集対象のIDが指定されていません');
  var month = Number(body.month);
  if (!month) throw new Error('編集対象の月が指定されていません');

  var ss = getCurrentYearSpreadsheet_();
  var sheetName = '不良' + month + '月';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('シート「' + sheetName + '」が見つかりません');

  var oldRows = findRecordRows_(sheet, body.id);
  if (oldRows.length === 0) throw new Error('編集対象の記録が見つかりませんでした(既に削除された可能性があります)');

  var mainRow = oldRows[0];
  var mainValues = sheet.getRange(mainRow, 1, 1, 24).getValues()[0];
  if (mainValues[2] !== verifiedName) throw new Error('自分が送信した記録のみ編集できます');
  var originalTimestamp = mainValues[0];

  var items = Array.isArray(body.items) ? body.items : [];
  if (items.length === 0) throw new Error('不良項目が指定されていません');

  // 元の行数より多くの行が必要な場合、直後の行が本当に空いているか確認する(他の記録を巻き込まないため)
  if (items.length > oldRows.length) {
    for (var r = mainRow + oldRows.length; r < mainRow + items.length; r++) {
      if (r > DATA_END_ROW) throw new Error(sheetName + ' が満杯です(不良項目の追加行)');
      var check = sheet.getRange(r, 1, 1, 14).getValues()[0];
      var hasData = check[0] !== '' || check[13] !== ''; // A列(タイムスタンプ) or N列(不良項目)
      if (hasData) throw new Error('不良項目を増やすための空き行が見つかりませんでした。項目数を減らすか、直接シートを編集してください。');
    }
  }

  // 元の行(枠線含む)を一旦すべてクリアしてから書き直す。
  // K・T・W列(良品数・金額・不良率)はSetupSpreadsheet.gs側の数式が全行に入っているため、
  // 巻き込んで消さないよう列を分けてクリアする(writeRecordRows_が書き込む列と同じ切り方)。
  var oldLastRow = mainRow + Math.max(oldRows.length, items.length) - 1;
  var clearRows = oldLastRow - mainRow + 1;
  sheet.getRange(mainRow, 1, clearRows, 10).clearContent();  // A-J
  sheet.getRange(mainRow, 12, clearRows, 7).clearContent();  // L-R
  sheet.getRange(mainRow, 19, clearRows, 1).clearContent();  // S
  sheet.getRange(mainRow, 21, clearRows, 2).clearContent();  // U-V
  sheet.getRange(mainRow, 24, clearRows, 1).clearContent();  // X
  sheet.getRange(mainRow, 1, clearRows, 24).setBorder(false, false, false, false, false, false);

  writeRecordRows_(sheet, mainRow, body, verifiedName, body.id, items, originalTimestamp);
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

  // --- 月次サマリー: 月別 個数・件数・金額(KP・差し戻し)、KP/差し戻しの年計個数 ---
  // 【2026-08-19改訂】差し戻しでも単価を入力するようになり「差し戻し 不良金額」行(8行目)が
  // 新設されたため、合計件数・合計個数の行が9・10行目から10・11行目へ1つずつ後ろにずれている。
  var summarySheet = ss.getSheetByName('月次サマリー');
  var monthlyQty = summarySheet.getRange(12, 2, 1, MONTHS.length).getValues()[0].map(Number);   // 合計 不良個数(KP+差し戻し)
  var monthlyCount = summarySheet.getRange(11, 2, 1, MONTHS.length).getValues()[0].map(Number);  // 合計 不良件数(KP+差し戻し)
  var monthlyAmountKP = summarySheet.getRange(4, 2, 1, MONTHS.length).getValues()[0].map(Number); // KP 不良金額
  var monthlyAmountRework = summarySheet.getRange(9, 2, 1, MONTHS.length).getValues()[0].map(Number); // 差し戻し 不良金額
  var monthlyQtyKP = summarySheet.getRange(3, 2, 1, MONTHS.length).getValues()[0].map(Number);   // KP 不良個数
  var monthlyQtyRework = summarySheet.getRange(8, 2, 1, MONTHS.length).getValues()[0].map(Number); // 差し戻し 不良個数
  var monthlyCountKP = summarySheet.getRange(2, 2, 1, MONTHS.length).getValues()[0].map(Number);   // KP 不良件数
  var monthlyCountRework = summarySheet.getRange(6, 2, 1, MONTHS.length).getValues()[0].map(Number); // 差し戻し 不良件数
  var kpQtyYear = Number(summarySheet.getRange(3, 14).getValue()) || 0;     // KP 不良個数(年計)
  var reworkQtyYear = Number(summarySheet.getRange(8, 14).getValue()) || 0; // 差し戻し 不良個数(年計)

  // --- 不良集計: キズ系項目ごとの月別件数・個数(SUMIF/COUNTIF相当をJS側で計算)。
  // ダッシュボードの「月別キズ不良件数」「月別キズ不良個数」グラフをキズ項目ごとに積み上げるため
  // (2026-08-19。以前あった分類別(5グループ)の内訳は、キズ系の内訳グラフと役割が重複するため削除した)。
  var itemSheet = ss.getSheetByName('不良集計');
  var itemNames = itemSheet.getRange(3, 1, DEFECT_ITEMS.length, 1).getValues().map(function (r) { return r[0]; });
  var kizuItemNames = DEFECT_ITEMS.filter(function (item) { return item.group === 'キズ系'; }).map(function (item) { return item.name; });
  var stackedByKizuItem = [];
  var stackedByKizuItemCount = [];
  MONTHS.forEach(function (month, mi) {
    var countCol = 3 + mi * 2; // 不良集計シートの月別「件数」列(C,E,G...)
    var qtyCol = countCol + 1; // 同じ月の「個数」列(D,F,H...)
    var countValues = itemSheet.getRange(3, countCol, DEFECT_ITEMS.length, 1).getValues().map(function (r) { return Number(r[0]) || 0; });
    var qtyValues = itemSheet.getRange(3, qtyCol, DEFECT_ITEMS.length, 1).getValues().map(function (r) { return Number(r[0]) || 0; });
    stackedByKizuItem.push(kizuItemNames.map(function (name) {
      var idx = itemNames.indexOf(name);
      return idx >= 0 ? qtyValues[idx] : 0;
    }));
    stackedByKizuItemCount.push(kizuItemNames.map(function (name) {
      var idx = itemNames.indexOf(name);
      return idx >= 0 ? countValues[idx] : 0;
    }));
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
  // 見出し文字列そのもの(「検査員」「検査員名」等)がデータ行に紛れ込んでいた場合に、実在の1人として
  // 誤カウントしないよう除外する(2026-08-17、実機で「検査員名」が集計に混入する不具合が判明したため対応)。
  var CC_LABEL_LIKE_VALUES_ = ['加工者', '加工者名', '検査員', '検査員名'];
  var kakoshaCount = {}, kensainCount = {};
  var ccSheetForRanking = ss.getSheetByName(CC_LEDGER_SHEET_NAME);
  if (ccSheetForRanking) {
    var ccCols = resolveCcLedgerColumns_(ccSheetForRanking);
    var ccLastRow = ccSheetForRanking.getLastRow();
    var ccRows = Math.max(ccLastRow - ccCols.dataStartRow + 1, 0);
    if (ccRows > 0 && ccCols.workerCol !== -1 && ccCols.inspectorCol !== -1) {
      var kakoshaValues = ccSheetForRanking.getRange(ccCols.dataStartRow, ccCols.workerCol, ccRows, 1).getValues();
      var kensainValues = ccSheetForRanking.getRange(ccCols.dataStartRow, ccCols.inspectorCol, ccRows, 1).getValues();
      kakoshaValues.forEach(function (r) {
        var v = r[0] ? r[0].toString().trim() : '';
        if (v && CC_LABEL_LIKE_VALUES_.indexOf(v) === -1) kakoshaCount[v] = (kakoshaCount[v] || 0) + 1;
      });
      kensainValues.forEach(function (r) {
        var v = r[0] ? r[0].toString().trim() : '';
        if (v && CC_LABEL_LIKE_VALUES_.indexOf(v) === -1) kensainCount[v] = (kensainCount[v] || 0) + 1;
      });
    }
  }

  // --- 不良〇月シート12枚: 得意先別金額・加工者別件数・月別加工数合計(不良率の分母) ---
  var customerAmount = {}; // { 得意先名: 金額合計 }
  var workerDefectCount = {}; // { 加工者: 不良件数(行数、追加行は加工者が空欄のため二重カウントされない) }
  var monthlyVolume = [];  // 月別 加工数合計(不良率の分母。ただし下記の通り一部除外あり)

  MONTHS.forEach(function (month) {
    var sheet = ss.getSheetByName('不良' + month + '月');
    var rows = sheet.getRange(2, 1, 114, 20).getValues(); // A〜T列(B処置区分/E得意先名/G加工者/J加工数/M不良数計/T金額)
    var volume = 0;
    rows.forEach(function (row) {
      var shochiKubun = row[1], customer = row[4], worker = row[6], suryo = row[9], totalDefectQty = row[12], amount = row[19];
      // 差し戻しで不良数計(M列)が0のレコードは不良率の分母(加工数)から除外する(2026-08-19、ユーザー要望)。
      // 差し戻しは「差し戻された時点で不良が発生している」ため、そのレコードの不良率が0%になるのは
      // 実態を正しく表しておらず、加工数だけを分母に加えると不良率(KP＋差し戻し合算)を不自然に薄めてしまうため。
      var excludeFromRate = (shochiKubun === '差し戻し' && Number(totalDefectQty) === 0);
      if (suryo && !excludeFromRate) volume += Number(suryo) || 0;
      if (worker) workerDefectCount[worker] = (workerDefectCount[worker] || 0) + 1;
      if (customer && amount) customerAmount[customer] = (customerAmount[customer] || 0) + (Number(amount) || 0);
    });
    monthlyVolume.push(volume);
  });

  var defectRate = MONTHS.map(function (month, mi) {
    var v = monthlyVolume[mi];
    return v > 0 ? Math.round((monthlyQty[mi] / v) * 1000) / 10 : 0;
  });

  // KPIタイル「社内不良(KP)比率(年計)」でのみ使う(2026-08-17、単独グラフ「社内不良(KP)と差し戻しの比率」は削除済み)
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
    monthlyAmountRework: monthlyAmountRework,
    monthlyQtyKP: monthlyQtyKP,
    monthlyQtyRework: monthlyQtyRework,
    monthlyCountKP: monthlyCountKP,
    monthlyCountRework: monthlyCountRework,
    kizuItems: kizuItemNames,
    stackedByKizuItem: stackedByKizuItem,
    stackedByKizuItemCount: stackedByKizuItemCount,
    causeGroups: causeGroups,
    causeTotals: causeTotals,
    customers: topN(customerAmount, 8),
    workers: topN(workerDefectCount, 8),
    defectRate: defectRate,
    kpReworkRatio: kpReworkRatio,
    claimMonthly: claimMonthly,
    claimByWorker: topN(kakoshaCount, 8),
    claimByInspector: topN(kensainCount, 8),
    updatedAt: new Date().toISOString()
  };
}

/** シートの次の空き行を探す(A列・N列のどちらかが埋まっていればその行は使用済みとみなす) */
function findNextRow_(sheet) {
  var rowCount = DATA_END_ROW - DATA_START_ROW + 1;
  var colA = sheet.getRange(DATA_START_ROW, 1, rowCount, 1).getValues();
  var colM = sheet.getRange(DATA_START_ROW, 14, rowCount, 1).getValues(); // N列(不良項目)
  var last = DATA_START_ROW - 1;
  for (var i = 0; i < colA.length; i++) {
    if (colA[i][0] !== '' || colM[i][0] !== '') last = DATA_START_ROW + i;
  }
  return last + 1;
}
