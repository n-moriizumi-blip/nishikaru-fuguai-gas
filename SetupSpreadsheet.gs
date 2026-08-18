/**
 * 品質不具合管理システム(新規構築プロジェクト) Phase 1
 * スプレッドシート構成の初期セットアップ用スクリプト
 *
 * 【実行方法】
 * 1. 対象スプレッドシート「客先クレーム(CC)・社内不良(KP)管理台帳 2026年度」を開く
 *    (https://docs.google.com/spreadsheets/d/1xeEybU6fhqvAEM02eEkUwBdOA_mD-fMUqKoCmckUcM0/edit)
 * 2. 拡張機能 > Apps Script を開き、このファイルの内容を貼り付ける
 * 3. 関数「setupQualityDefectSystem」を選択して1回だけ実行する
 * 4. 実行後、デフォルトで作られる「シート1」は不要なら手動で削除してよい
 *
 * このスクリプトは実行するたびに既存シートを作り直す(初期セットアップ専用)。
 * 本運用開始後に再実行すると入力済みデータが消えるので注意。
 *
 * 【2026-08-02改訂】社内不良(KP)・差し戻しを別シートにしていたが、項目がほぼ重複していたため
 * 「不良〇月」1本のシートに統合した(処置区分列で区別)。
 * 【2026-08-11改訂】「数量」を「加工数」に改名(何の数量か分かりにくいとの指摘)。「良品数」は
 * 手入力ではなく「加工数−不良数計」の自動計算式に変更。あわせて見やすさ向上のため、ヘッダー色・
 * 縞模様・処置区分による行の色分け(背景色+文字色)・集計シートの分類別色分け・グリッド罫線・
 * ヘッダー下線・グループ区切り線を追加した。
 * 【2026-08-11改訂その2】既存の「改善計画書台帳」(客先クレーム(CC)・社内不良(品証)(KP)・
 * 社内不良(製造工程)(SK)の3種が1枚のシートに縦積みで読みにくかった)を、独立した3シートに
 * 分けて追加した(不良〇月とは別物、不適合改善計画書のなぜなぜ分析〜水平展開までの進捗管理用)。
 * 台帳番号は自動採番せず、これまで通り手入力。月別分割はせず通し番号の1枚のシートのまま。
 * 【2026-08-13改訂】ファイル名を旧システムと同じ命名規則(「客先クレーム(CC)・社内不良(KP)管理台帳
 * 〇〇年度」)に変更(1回だけ`renameToLegacyNaming`を実行、詳細は同関数のコメント参照)。
 * あわせて年度自動ロールオーバーに対応: `getCurrentYearSpreadsheet_`が今日の日付から年度を判定し、
 * その年度のスプレッドシートを名前で検索、無ければ新規作成(setupQualityDefectSystemFor_で組み立て)
 * する。`setupQualityDefectSystem`以外の手動実行系関数・WebApi.gsは全てこの関数経由で対象
 * スプレッドシートを取得するため、6月になって新年度に入ると次のアクセス時に自動的に新しい
 * スプレッドシートへ切り替わる(旧コード.jsのgetOrCreateYearSpreadsheetと同じ考え方)。
 */

var SPREADSHEET_ID = '1xeEybU6fhqvAEM02eEkUwBdOA_mD-fMUqKoCmckUcM0'; // 2026年度分。年度が変わったらgetCurrentYearSpreadsheet_で別IDに自動で切り替わる

// 年度ごとのファイル名の接頭辞(末尾に「2026年度」等が付く、末尾スペース含む。旧システムと同じ命名規則、2026-08-13〜)
var SPREADSHEET_NAME_PREFIX = '客先クレーム(CC)・社内不良(KP)管理台帳 ';
// 年度スプレッドシートの検索・新規作成先フォルダ。
// 【2026-08-13訂正】「品質不具合管理システム(改訂中)」フォルダ(1cT2gttoIR69MYn2KRV3_E3lv0mjxhVeo)を
// 指定していたが、実際の2026年度ファイルはその1つ上の「改善計画書台帳」フォルダ直下にあり、
// 検索が常に「見つからない」と判定して重複スプレッドシートを自動作成し続けるバグになっていた。
// 実際のファイル置き場に合わせてこちらのフォルダIDに修正した。
var SPREADSHEET_PARENT_FOLDER_ID = '1zWL_FZ_yMjx5pTDyK6fdQf0Tdkf-wvrC';

// 組織図マスタ(機械名・加工者名のプルダウン用、WebApi.gsのORG_MASTER_SS_IDと同じ)
var ORG_MASTER_SS_ID = '1fffjE_bwrzswvRO62U0OHwvqrs5b_UuSV5IbudUMxec';
var ORG_MASTER_SHEET_NAME = 'プルダウン用';

var MONTHS = [6, 7, 8, 9, 10, 11, 12, 1, 2, 3, 4, 5]; // 6月始まりの年度

// 処置区分(社内不良(KP)として確定させたか、前工程へ差し戻したか。1件の不良につきどちらか一方)
var SHOCHI_KUBUN = ['社内不良(KP)', '差し戻し'];

// 見やすさ向上のための色設定
var COLOR = {
  HEADER_BG: '#1F3A52',
  HEADER_FONT: '#FFFFFF',
  HEADER_BORDER: '#0F1D28',
  BAND_ALT: '#F4F6F8',
  GRID_BORDER: '#D5DBE0',
  KP_BG: '#DCE9F5',
  KP_FONT: '#1F3A52',
  REWORK_BG: '#FDECC8',
  REWORK_FONT: '#7A5A00',
  SUMMARY_KP_BG: '#F5F9FC',
  SUMMARY_REWORK_BG: '#FFFBF2',
  SUMMARY_TOTAL_BG: '#E4EAF0',
  TOTAL_FONT: '#1F3A52',
  SUMMARY_COUNT_BG: '#EEF3F8', // 不良集計等の「件数」列見出し(同系色の薄い方)
  SUMMARY_QTY_BG: '#D7E3F0'    // 不良集計等の「個数」列見出し(同系色の濃い方)
};

// 不良集計シートの「分類」ごとの色(登場順に自動で割り当てる)
var GROUP_PALETTE = ['#E8F0FE', '#FDEEE0', '#EAF7EA', '#FBEAF5', '#FFF6D9', '#E6F7F5', '#F0E9FB', '#F0F0F0'];

// 不良項目マスタ(40項目、2026-08-12に「客先クレーム(CC)・社内不良(KP)管理台帳 2026年度」の
// 「社内不具合〇月」シートG列(不具合内容1)の実際のプルダウン(データ!G4:G44、12ヶ月とも同一)に
// そのまま合わせて全面差し替え。以前の29項目マスタ(差し戻し品一覧表ベース)は廃止した。
// グループ分けは元のプルダウンの並び順から見た区切りをもとにした案(ユーザー承認済み)。
var DEFECT_ITEMS = [
  // 寸法系
  { group: '寸法系', name: '寸法不良' },
  { group: '寸法系', name: '寸法大' },
  { group: '寸法系', name: '寸法小' },
  { group: '寸法系', name: '外径大' },
  { group: '寸法系', name: '外径小' },
  { group: '寸法系', name: '内径大' },
  { group: '寸法系', name: '内径小' },
  { group: '寸法系', name: '長さ不良' },
  { group: '寸法系', name: 'ねじ不良' },
  { group: '寸法系', name: '形状不良' },
  { group: '寸法系', name: '穴ズレ（穴曲り）' },
  { group: '寸法系', name: 'テーパ不良' },
  { group: '寸法系', name: '平研不良' },
  // 変形系
  { group: '変形系', name: '曲がり' },
  { group: '変形系', name: '変形' },
  { group: '変形系', name: '振れ大' },
  // 材質・混入系
  { group: '材質・混入系', name: '未加工混入（未加工）' },
  { group: '材質・混入系', name: '材質違い' },
  { group: '材質・混入系', name: '圧入・接着不良' },
  { group: '材質・混入系', name: '現品違い' },
  // キズ系
  { group: 'キズ系', name: 'キズ' },
  { group: 'キズ系', name: '流動キズ' },
  { group: 'キズ系', name: 'チャックキズ' },
  { group: 'キズ系', name: '回収時キズ' },
  { group: 'キズ系', name: '落下キズ' },
  { group: 'キズ系', name: '材料キズ' },
  { group: 'キズ系', name: '加工時キズ' },
  { group: 'キズ系', name: 'ガイドブッシュキズ' },
  { group: 'キズ系', name: 'その他キズ' },
  // 外観・その他系
  { group: '外観・その他系', name: 'バリ' },
  { group: '外観・その他系', name: '段差' },
  { group: '外観・その他系', name: '溶着' },
  { group: '外観・その他系', name: '挽目不良' },
  { group: '外観・その他系', name: 'バフがけ不良' },
  { group: '外観・その他系', name: '修正不良' },
  { group: '外観・その他系', name: '取り残し' },
  { group: '外観・その他系', name: 'ムシレ' },
  { group: '外観・その他系', name: '切粉' },
  { group: '外観・その他系', name: '汚れ' },
  { group: '外観・その他系', name: '変色・錆' }
];

// 「クレーム集計」シート・ダッシュボードの加工者別/検査員別集計の集計元となる
// 「客先クレーム管理台帳(CC)」の列位置。
// 【2026-08-18訂正】旧デザイン(結合セルの複数行ヘッダー)の実機調査(2026-08-17)で得た列番号を
// 一時的に固定値で持っていたが、buildLedgerSheetV2_導入(台帳のリニューアル)で列位置が変わるため、
// resolveCcLedgerColumns_で見出し文字列から動的に求める方式に変更した(旧デザイン・新デザインの
// どちらが実際に稼働していても、見出し名さえ一致すれば自動的に追随する)。
var CC_LEDGER_SHEET_NAME = '客先クレーム管理台帳(CC)';
var CC_DATA_END_ROW = 1000; // 手入力で増えていく台帳のため、余裕を持った行数まで集計対象にする

/**
 * 「客先クレーム管理台帳(CC)」の見出し行(固定行数の範囲)を実際に読み取り、発生日・クレーム内容分類・
 * 加工者・検査員の列番号とデータ開始行(固定行数+1)を求める(2026-08-18新設)。見出しセルの装飾用
 * 空白詰めを除去してから部分一致で探すため、台帳のデザインが変わっても見出し文字列さえ同じなら
 * 追随できる(旧デザイン=結合セルの複数行ヘッダー、新デザイン=buildLedgerSheetV2_の2行ヘッダー、
 * どちらでも同じロジックで動作することを確認済み)。
 */
function resolveCcLedgerColumns_(sheet) {
  var frozen = sheet.getFrozenRows() || 1;
  var lastCol = sheet.getLastColumn();
  var headerRows = sheet.getRange(1, 1, frozen, lastCol).getValues();

  function findCol(name) {
    for (var r = 0; r < headerRows.length; r++) {
      for (var c = 0; c < headerRows[r].length; c++) {
        var text = (headerRows[r][c] || '').toString().replace(/[\s　]+/g, '');
        if (text.indexOf(name) !== -1) return c + 1;
      }
    }
    return -1;
  }

  return {
    dateCol: findCol('発生日'),
    categoryCol: findCol('クレーム内容分類'),
    workerCol: findCol('加工者'),
    inspectorCol: findCol('検査員'),
    dataStartRow: frozen + 1
  };
}

// 「社内不良管理台帳(製造工程)(SK)」の発生日列。2026-08-17、debugInspectDateColumnValidationの実機調査で
// C列(発生日(品質会議開催日))だけ日付の入力規則が無く表示形式も壊れている(0.###############)ことが
// 判明したため、CC/KPと同じ日付検証+表示形式に直す(fixSkDateColumnValidation)。
var SK_LEDGER_SHEET_NAME = '社内不良管理台帳(製造工程)(SK)';
var SK_DATE_COL = 3;      // C列(発生日(品質会議開催日))
var SK_DATA_START_ROW = 5; // 固定行数4(ヘッダーが1〜4行目)のため、データは5行目から

// キズ原因マスタ(現行「キズ集計管理台帳」A〜Hの実データから踏襲。任意項目、主にキズ系の不良で原因分析用に使う)
var KP_CAUSE_ITEMS = [
  { code: 'A', group: '流動', name: '洗浄時' },
  { code: 'A', group: '流動', name: 'ざるに入れてエアー吹き' },
  { code: 'A', group: '流動', name: '製品の移し替え時' },
  { code: 'A', group: '流動', name: '箱の中での製品同士の打痕' },
  { code: 'A', group: '流動', name: '落下' },
  { code: 'B', group: 'チャック', name: 'チャック内の切粉による圧痕' },
  { code: 'B', group: 'チャック', name: '切粉が絡んでのチャック' },
  { code: 'B', group: 'チャック', name: 'ワーク掴み方' },
  { code: 'C', group: '回収時', name: '回収時のワークの落ち方' },
  { code: 'C', group: '回収時', name: '自動回収機のワークがいっぱい' },
  { code: 'C', group: '回収時', name: 'NC回収時の落下' },
  { code: 'C', group: '回収時', name: 'NC回収時の打痕' },
  { code: 'D', group: '落下', name: '突っ切り落とし加工時' },
  { code: 'D', group: '落下', name: '落下' },
  { code: 'E', group: '材料', name: '材料のキズ' },
  { code: 'E', group: '材料', name: '材料の面取り時の圧痕' },
  { code: 'F', group: '加工時', name: '切粉がワークに絡まってのキズ' },
  { code: 'G', group: 'ガイドブッシュ', name: 'ガイドブッシュのキズ(外径挽かない)' },
  { code: 'H', group: 'その他', name: 'その他' }
];

// 不適合改善計画書ワークフロー欄(CC・KP(品証)・SK(製造工程)の3台帳で共通)。
// 【2026-08-18全面リニューアル】旧デザイン(WORKFLOW_HEADERS、月/日プルダウン2列+⇒矢印列)は廃止。
// ユーザーが旧システム形式で手作りした本物の3台帳をxlsxダウンロードで実機調査した結果、
// ①ワークフロー完了日が「⇒/月/日」の3列プルダウン(年の記録なし)だった、②「効果確認」は
// 加工者・品証・社長の3部署がそれぞれ確認する構成だった、と判明(2026-08-17〜18)。
// これを踏まえ、完了日は発生日と同じ単一の日付セルに、効果確認は3ステップに分けて再設計した
// (buildLedgerSheetV2_参照)。詳細な調査結果・合意事項はCLAUDE.md参照。
var LEDGER_WORKFLOW_STEPS = [
  'なぜなぜ分析', '改善策の共有', '図面に貼付', '改善の実施',
  '効果確認(加工者)', '効果確認(品証)', '効果確認(社長)',
  '水平展開の実施', '水平展開の効果確認'
];
var HANKO_CHOICES_ = ['〇', '×'];
var SHUKKA_KUBUN_CHOICES_ = ['今回製作品', '在庫出荷品'];

/**
 * このスプレッドシート自身の「データ」シートから、ワークフロー欄の「担当」用の全社員名簿(C列)と
 * 検査員名簿(I列)を読み取る(2026-08-18新設)。機械名・加工者名は既存通り組織図マスタ
 * (fetchOrgMasterLists_)を正とするが、「担当」は品証・社長など加工者に限らない全社員が対象になり
 * うる上、検査員も組織図マスタに項目が無いため、ユーザーが実運用している「データ」シートの
 * この2列を正とする。「検査員名」等、見出しらしき文字列が値として紛れていた場合は除外する
 * (WebApi.gsのCC_LABEL_LIKE_VALUES_と同じ対策、実機で混入を確認済み)。
 */
function fetchDataSheetLists_(ss) {
  var sheet = ss.getSheetByName('データ');
  if (!sheet) return { staff: [], inspectors: [] };
  var labelLike = ['担当者', '担当者名', '検査員', '検査員名', '加工者', '加工者名'];
  function uniqueNonBlank(range) {
    var seen = [], result = [];
    range.getValues().forEach(function (r) {
      var v = r[0] ? r[0].toString().trim() : '';
      if (v && labelLike.indexOf(v) === -1 && seen.indexOf(v) === -1) { seen.push(v); result.push(v); }
    });
    return result;
  }
  var lastRow = Math.max(sheet.getLastRow(), 2);
  return {
    staff: uniqueNonBlank(sheet.getRange(2, 3, lastRow - 1, 1)),      // C列
    inspectors: uniqueNonBlank(sheet.getRange(2, 9, lastRow - 1, 1))  // I列
  };
}

function setupQualityDefectSystem() {
  setupQualityDefectSystemFor_(SpreadsheetApp.openById(SPREADSHEET_ID));
}

/**
 * setupQualityDefectSystemの本体。対象スプレッドシートを引数で受け取る形にして、
 * 年度自動ロールオーバー(getOrCreateYearSpreadsheet_)から新規スプレッドシートの
 * 組み立てにも使い回せるようにしてある(2026-08-13)。
 */
function setupQualityDefectSystemFor_(ss) {
  MONTHS.forEach(function (month) {
    buildDefectMonthlySheet_(ss, month);
  });

  buildDefectSummarySheet_(ss);
  buildKpCauseSummarySheet_(ss);
  buildMonthlySummarySheet_(ss);
  buildClaimSummarySheet_(ss);
  buildDefectDashboardSheet_(ss);

  var masters = fetchOrgMasterLists_();
  var dataLists = fetchDataSheetLists_(ss);
  buildCcLedgerSheetV2_(ss, masters, dataLists, '客先クレーム管理台帳(CC)仮').hideSheet();
  buildKpLedgerSheetV2_(ss, masters, dataLists, '社内不良管理台帳(品証)(KP)(仮)').hideSheet();
  buildSkLedgerSheetV2_(ss, masters, dataLists, '社内不良管理台帳(製造工程)(SK)(仮)').hideSheet();

  SpreadsheetApp.flush();
  Logger.log('セットアップ完了(' + ss.getName() + '): 不良12シート、集計4シート、グラフ1シート、改善計画書台帳3シートを作成しました。');
}

/**
 * 年度自動ロールオーバー(2026-08-13新設)。6月始まりの年度ごとに
 * 「客先クレーム(CC)・社内不良(KP)管理台帳 〇〇年度」というスプレッドシートを自動的に探し、
 * 無ければ新規作成する(旧コード.jsのgetOrCreateYearSpreadsheetと同じ考え方)。
 * 新規作成時はDrive上に空のスプレッドシートを作ってから品質不具合管理システムフォルダへ移動し、
 * setupQualityDefectSystemFor_で不良〇月12シート・集計3シート・グラフ・改善計画書台帳3シートを
 * 組み立てる。作成直後の既定シート(シート1/Sheet1)は組み立て後に削除する。
 *
 * 【2026-08-13訂正】検索範囲をDrive全体(DriveApp.getFilesByName)からSPREADSHEET_PARENT_FOLDER_ID
 * フォルダ直下だけに変更した。旧システムの退役済みファイル(改善計画書台帳フォルダ直下、
 * 品質不具合管理システムフォルダの外)が同名になっても検索に引っかからず、旧ファイルを一切
 * 触らずに済むようにするため。旧ファイルは編集権限が無く(共有ドライブ内、別権限设定)
 * renameToLegacyNamingで名前変更しようとして「Access denied: DriveApp」で失敗したことが判明した。
 */
function getCurrentFiscalYear_(date) {
  date = date || new Date();
  var month = date.getMonth() + 1;
  return month >= 6 ? date.getFullYear() : date.getFullYear() - 1;
}

function getOrCreateYearSpreadsheet_(fiscalYear) {
  var fileName = SPREADSHEET_NAME_PREFIX + fiscalYear + '年度';
  var folder = DriveApp.getFolderById(SPREADSHEET_PARENT_FOLDER_ID);
  var files = folder.getFilesByName(fileName); // フォルダ直下だけを検索(Drive全体は検索しない)
  if (files.hasNext()) return SpreadsheetApp.open(files.next());

  Logger.log('「' + fileName + '」が見つからないため新規作成します。');
  var newSs = SpreadsheetApp.create(fileName);
  var newFile = DriveApp.getFileById(newSs.getId());
  folder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile); // create()はマイドライブ直下に作られるため、フォルダへ移動する

  var defaultSheetNames = newSs.getSheets().map(function (s) { return s.getName(); });
  setupQualityDefectSystemFor_(newSs);
  defaultSheetNames.forEach(function (name) {
    var s = newSs.getSheetByName(name);
    if (s && newSs.getSheets().length > 1) newSs.deleteSheet(s);
  });

  // 前年度のCC/KP/SK台帳(手作業で旧システム形式に作り直したもの)・データ・サプライヤー不具合シートを
  // 構造・プルダウンだけ引き継ぐ(2026-08-15、ユーザー合意)。前年度が無ければ諦めて空のまま進める。
  try {
    var prevSs = getOrCreateYearSpreadsheet_(fiscalYear - 1);
    copyCurrentStructureSheets_(prevSs, newSs);
  } catch (err) {
    Logger.log('前年度シートの引き継ぎに失敗しました(手動での確認が必要): ' + err.message);
  }

  return newSs;
}

/**
 * 新年度作成時に「構造・プルダウンだけ引き継ぎ、入力済みデータは空にする」対象のシート(2026-08-15新設)。
 * 客先クレーム管理台帳(CC)・社内不良管理台帳(品証)(KP)・社内不良管理台帳(製造工程)(SK)は、
 * ユーザーが2026-08-15に旧システム形式(結合セルを含む複数行ヘッダー)で手作業で作り直したため、
 * コードで再現せずシートをそのままコピーする方式にしている。
 * 「データ」は年月日・不良項目名等のプルダウン用マスタ(年度が変わっても内容ごと引き継ぐべき参照データ)
 * のため、clearDataをfalseにして中身もそのままコピーする。
 * 「サプライヤー不具合」は固定行数(フリーズ)が設定されていないため、見出し行数を2行(タイトル行＋
 * 項目名行)に決め打ちしている(2026-08-15時点の実際の構成から確認した値)。手動で構成を変えた場合は
 * この値も見直すこと。
 */
var COPY_STRUCTURE_SHEETS_ = [
  { name: '客先クレーム管理台帳(CC)', clearData: true },
  { name: '社内不良管理台帳(品証)(KP)', clearData: true },
  { name: '社内不良管理台帳(製造工程)(SK)', clearData: true },
  { name: 'データ', clearData: false },
  { name: 'サプライヤー不具合', clearData: true, headerRowsOverride: 2 }
];

function copyCurrentStructureSheets_(sourceSs, targetSs) {
  var log = [];
  COPY_STRUCTURE_SHEETS_.forEach(function (cfg) {
    var source = sourceSs.getSheetByName(cfg.name);
    if (!source) { log.push(cfg.name + ': コピー元に見つからないためスキップ'); return; }
    if (targetSs.getSheetByName(cfg.name)) { log.push(cfg.name + ': 既にあるためスキップ'); return; }

    var copied = source.copyTo(targetSs);
    copied.setName(cfg.name); // copyToは「〇〇のコピー」という名前になるため、元の名前に揃える
    if (source.isSheetHidden()) copied.hideSheet();

    if (cfg.clearData) {
      var headerRows = cfg.headerRowsOverride || source.getFrozenRows();
      var lastRow = copied.getMaxRows();
      var lastCol = copied.getMaxColumns();
      if (headerRows > 0 && lastRow > headerRows) {
        copied.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol).clearContent();
      }
    }
    log.push(cfg.name + ': コピーしました' + (cfg.clearData ? '(見出し行より下のデータはクリア)' : '(内容もそのまま引き継ぎ)'));
  });
  Logger.log(log.join('\n'));
}

/** 現在の年度(今日の日付基準)のスプレッドシートを返す(無ければ自動作成)。dateを省略すると今日の日付を使う。 */
function getCurrentYearSpreadsheet_(date) {
  return getOrCreateYearSpreadsheet_(getCurrentFiscalYear_(date));
}

/**
 * 【1回だけ手動実行】このシステムのスプレッドシート名を、旧システムと同じ命名規則
 * 「客先クレーム(CC)・社内不良(KP)管理台帳 2026年度」に変更する(2026-08-13)。
 * 【2026-08-13訂正】旧システムの2026年度ファイル(退役済み)は編集権限が無く名前変更できなかった
 * ため、そちらは変更しない方針にした(getOrCreateYearSpreadsheet_の検索範囲をフォルダ内だけに
 * 絞ったことで、名前が重複していても支障が無くなったため、無理に触る必要が無くなった)。
 * IDは変わらないため、このリネーム自体はSetupSpreadsheet.gs/WebApi.gsのコードに影響しない。
 */
function renameToLegacyNaming() {
  var newName = SPREADSHEET_NAME_PREFIX + getCurrentFiscalYear_() + '年度';
  DriveApp.getFileById(SPREADSHEET_ID).setName(newName);
  Logger.log('新システムのファイル名を「' + newName + '」に変更しました。');
}

/**
 * 【手動実行用・年1回】次の年度のスプレッドシートを先に準備しておく(2026-08-15新設)。
 * 通常は年度が変わった直後に最初にアクセスした人(ダッシュボード表示 or 不具合入力)が
 * 自動作成のきっかけになり、setupQualityDefectSystemFor_(不良12シート等の組み立て、前年度の
 * CC/KP/SK台帳等のコピー)に数分かかって重く感じられてしまう。5月末など年度が変わる少し前に
 * この関数を実行しておけば、その重い処理を事前に済ませておけるため、6月1日以降の実際の
 * アクセスはいつも通りの速さで動く。既に来年度のファイルがある場合は何もしない(毎年5月末に
 * 実行する運用でよく、何度実行しても安全)。
 */
function prepareNextYearSpreadsheet() {
  var nextYear = getCurrentFiscalYear_() + 1;
  var ss = getOrCreateYearSpreadsheet_(nextYear);
  Logger.log('次年度(' + nextYear + '年度)のスプレッドシートを準備しました: ' + ss.getUrl());
}

/**
 * トリガー用ハンドラ(2026-08-15新設)。GASの時間主導型トリガーには「年1回」という設定が無いため、
 * 毎月25日に実行されるトリガーを設置し、この関数側で「5月かどうか」を判定して年1回に絞る方式にした
 * (6月の年度切り替わりの少し前、5月に1回だけ実際にprepareNextYearSpreadsheetを実行する)。
 */
function autoPrepareNextYearSpreadsheet() {
  var month = new Date().getMonth() + 1;
  if (month !== 5) return; // 5月以外は何もしない
  prepareNextYearSpreadsheet();
}

/**
 * 【手動実行用・1回だけ】毎月25日午前3時台にautoPrepareNextYearSpreadsheetを自動実行するトリガーを
 * 設置する(2026-08-15新設)。これを1回実行しておけば、以後は`prepareNextYearSpreadsheet`を毎年
 * 手動実行する必要が無くなる(5月に自動で次年度分の準備が走る)。既に同じ関数のトリガーがあれば
 * 重複作成しない(何度実行しても安全)。
 * 【注意】このプロジェクトで初めてトリガーを作成する場合、実行時に「承認が必要です」の画面が
 * 出ることがある(トリガー作成用の権限を新たに使うため)。出たら許可すること。
 */
function installPrepareNextYearTrigger() {
  var already = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'autoPrepareNextYearSpreadsheet';
  });
  if (already) {
    Logger.log('既にトリガーが設定されています(重複作成はしません)。');
    return;
  }
  ScriptApp.newTrigger('autoPrepareNextYearSpreadsheet')
    .timeBased()
    .onMonthDay(25)
    .atHour(3)
    .create();
  Logger.log('毎月25日の午前3時台に自動チェックするトリガーを設置しました。実際に次年度分を準備するのは5月だけです。');
}

/**
 * 【手動実行用・診断】設置されているトリガーの一覧を確認する(2026-08-15新設)。
 */
function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    Logger.log('トリガーは1つも設置されていません。');
    return;
  }
  var log = triggers.map(function (t, i) {
    return (i + 1) + '. 関数「' + t.getHandlerFunction() + '」 種類=' + t.getEventType() + ' トリガー元=' + t.getTriggerSource();
  });
  Logger.log(log.join('\n'));
}

/**
 * 【手動実行用】installPrepareNextYearTriggerで設置したトリガーを削除する(2026-08-15新設)。
 */
function removePrepareNextYearTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'autoPrepareNextYearSpreadsheet') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log(removed + '件のトリガーを削除しました。');
}

/**
 * 【手動実行用】DEFECT_ITEMSマスタ(不良項目)を更新した後に実行する。
 * setupQualityDefectSystemと違い、既存の「不良〇月」シートのデータ(6月・7月の移行分など)は
 * 一切消さず、M列のプルダウンの選択肢だけを新しいDEFECT_ITEMSに差し替える。
 * 「不良集計」「改善計画書台帳」3シートは中身が全て数式・空のプルダウンだけで実データが無いため、
 * 作り直して問題ない。
 */
function updateDefectItemMaster() {
  var ss = getCurrentYearSpreadsheet_();
  var itemNames = DEFECT_ITEMS.map(function (item) { return item.name; });
  var itemRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(itemNames, true)
    .setAllowInvalid(false)
    .build();

  MONTHS.forEach(function (month) {
    var sheet = ss.getSheetByName('不良' + month + '月');
    if (!sheet) return;
    sheet.getRange(2, 13, 114, 1).setDataValidation(itemRule); // M列(不良項目)
  });

  buildDefectSummarySheet_(ss);     // 実データが無い集計表なので作り直して問題ない
  buildKpCauseSummarySheet_(ss);    // 件数・個数の2列レイアウト(2026-08-12)を反映するため合わせて作り直す
  buildDefectDashboardSheet_(ss);   // グラフの元データが上記2シートの構造に依存するため合わせて作り直す

  var masters = fetchOrgMasterLists_();
  var dataLists = fetchDataSheetLists_(ss);
  buildCcLedgerSheetV2_(ss, masters, dataLists, '客先クレーム管理台帳(CC)仮').hideSheet();
  buildKpLedgerSheetV2_(ss, masters, dataLists, '社内不良管理台帳(品証)(KP)(仮)').hideSheet();
  buildSkLedgerSheetV2_(ss, masters, dataLists, '社内不良管理台帳(製造工程)(SK)(仮)').hideSheet();

  SpreadsheetApp.flush();
  Logger.log('不良項目マスタを更新しました(' + itemNames.length + '項目)。「不良〇月」12シートのM列プルダウンを更新、'
    + '「不良集計」「不良集計(キズ原因)」「改善計画書台帳」4シートは作り直しました(データへの影響なし)。');
}

/**
 * 【手動実行用】「不良7月」で手動調整した列幅を、他の11ヶ月の「不良〇月」シートにも揃える。
 * あわせて全シート・全セルを水平・垂直とも中央揃えにする。データは一切変更しない(書式のみ)。
 */
function syncMonthlySheetLayout() {
  var ss = getCurrentYearSpreadsheet_();
  var referenceSheet = ss.getSheetByName('不良7月');
  if (!referenceSheet) {
    Logger.log('【エラー】基準となる「不良7月」シートが見つかりません。');
    return;
  }
  var lastCol = referenceSheet.getLastColumn();
  var widths = [];
  for (var c = 1; c <= lastCol; c++) widths.push(referenceSheet.getColumnWidth(c));

  var updated = [];
  MONTHS.forEach(function (month) {
    var sheet = ss.getSheetByName('不良' + month + '月');
    if (!sheet) return;
    for (var i = 0; i < widths.length; i++) sheet.setColumnWidth(i + 1, widths[i]);
    sheet.getRange(1, 1, 115, lastCol).setHorizontalAlignment('center').setVerticalAlignment('middle');
    updated.push(sheet.getName());
  });

  Logger.log('「不良7月」の列幅(' + lastCol + '列)を基準に、以下のシートを更新しました(中央揃えも適用):\n' + updated.join('、'));
}

/** 組織図マスタから機械名・加工者名の一覧を取得(プルダウン用、WebApi.gsのgetMasters_と同じロジック) */
function fetchOrgMasterLists_() {
  var ss = SpreadsheetApp.openById(ORG_MASTER_SS_ID);
  var sheet = ss.getSheetByName(ORG_MASTER_SHEET_NAME);
  if (!sheet) return { kishu: [], kakosha: [] };

  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var colKishu = header.indexOf('機械名');
  var colKakosha = header.indexOf('加工者名');
  if (colKishu === -1 || colKakosha === -1) return { kishu: [], kakosha: [] };

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
 * 【手動実行用・既存データへの一括適用】全12ヶ月の「不良〇月」シートで、既に入力済みのレコード
 * (不良項目が複数で複数行にまたがる場合はそのまとまり)を1件ずつ枠線で囲む。
 * WebApi.gsのwriteDefectRecord_が新規送信のたびに行うようになった枠線付けを、
 * それ以前に入力済みだったデータ(6月・7月の移行分、テスト送信分など)にさかのぼって適用するための関数。
 * 1回実行すれば十分(何度実行しても結果は同じ)。
 */
function addRecordBordersToAllMonths() {
  var ss = getCurrentYearSpreadsheet_();
  var log = [];
  var startRow = 2;
  var endRow = 115;

  MONTHS.forEach(function (month) {
    var sheet = ss.getSheetByName('不良' + month + '月');
    if (!sheet) return;

    var rows = endRow - startRow + 1;
    var colA = sheet.getRange(startRow, 1, rows, 1).getValues();
    var colM = sheet.getRange(startRow, 13, rows, 1).getValues();

    var recordCount = 0;
    var groupStart = -1;
    for (var i = 0; i < rows; i++) {
      var rowNum = startRow + i;
      var hasA = colA[i][0] !== '';
      var hasM = colM[i][0] !== '';

      if (hasA) {
        // 新しいレコードの開始。直前のグループがあれば先に枠線を付ける
        if (groupStart !== -1) {
          sheet.getRange(groupStart, 1, rowNum - groupStart, 23)
            .setBorder(true, true, true, true, false, false, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
          recordCount++;
        }
        groupStart = rowNum;
      } else if (!hasM && groupStart !== -1) {
        // A・Mとも空欄の行が現れたらレコードの終わり
        sheet.getRange(groupStart, 1, rowNum - groupStart, 23)
          .setBorder(true, true, true, true, false, false, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        recordCount++;
        groupStart = -1;
      }
    }
    if (groupStart !== -1) {
      // シート末尾までレコードが続いていた場合
      sheet.getRange(groupStart, 1, startRow + rows - groupStart, 23)
        .setBorder(true, true, true, true, false, false, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      recordCount++;
    }
    log.push('不良' + month + '月: ' + recordCount + '件に枠線を付けました');
  });

  Logger.log(log.join('\n'));
}

/**
 * 【1回だけ手動実行・2026-08-13】既存の「不良〇月」12シートに「製造番号」列(D列、得意先名の左)を
 * 追加する。列構成をコード側(buildDefectMonthlySheet_)で変更したのに合わせて、入力済みデータがある
 * 既存スプレッドシートにも反映するための移行用関数。
 * sheet.insertColumnBefore()はデータ・数式・入力規則・列の非表示状態・条件付き書式の範囲を
 * 保持したまま列を1つ右にずらす(replaceSheet_で作り直す他の関数と違い、既存データは一切失われない)。
 * 「不良集計」等の集計シート側の数式も、列がずれたことをGoogle Sheetsが自動的に追従するため、
 * この関数の実行後に改めてbuildDefectSummarySheet_等を作り直す必要はない。
 * 何度実行しても、既に製造番号列がある場合はその月をスキップする(冪等)。
 */
function addManufacturingNumberColumn() {
  var ss = getCurrentYearSpreadsheet_();
  var log = [];
  MONTHS.forEach(function (month) {
    var sheet = ss.getSheetByName('不良' + month + '月');
    if (!sheet) return;
    if (sheet.getRange(1, 4).getValue() === '製造番号') {
      log.push('不良' + month + '月: 既に製造番号列があるためスキップ');
      return;
    }
    sheet.insertColumnBefore(4);
    sheet.getRange(1, 4).setValue('製造番号')
      .setFontWeight('bold').setBackground(COLOR.HEADER_BG).setFontColor(COLOR.HEADER_FONT)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setColumnWidth(4, 140);

    var bandColors = [];
    for (var r = 0; r < 114; r++) bandColors.push([(r % 2 === 0) ? '#FFFFFF' : COLOR.BAND_ALT]);
    sheet.getRange(2, 4, 114, 1).setBackgrounds(bandColors)
      .setVerticalAlignment('middle').setHorizontalAlignment('center');
    log.push('不良' + month + '月: 製造番号列を追加しました');
  });
  SpreadsheetApp.flush();
  Logger.log(log.join('\n'));
}

/**
 * 【診断用・手動実行】現在の年度のスプレッドシートに実際にあるシートを一覧表示する(2026-08-15新設)。
 * ユーザーが直接シートを追加・変更した場合に、コード側(setupQualityDefectSystemFor_等)が把握している
 * 構成とのズレを確認するために使う。シート名・表示/非表示・行数・列数をログに出す。
 */
function listAllSheets() {
  var ss = getCurrentYearSpreadsheet_();
  var log = ['スプレッドシート: ' + ss.getName()];
  ss.getSheets().forEach(function (sheet, i) {
    var lastRowWithData = sheet.getLastRow();
    log.push(
      (i + 1) + '. 「' + sheet.getName() + '」' +
      (sheet.isSheetHidden() ? '(非表示)' : '(表示)') +
      ' 行数=' + sheet.getMaxRows() + ' 列数=' + sheet.getMaxColumns() +
      ' 固定行数=' + sheet.getFrozenRows() +
      ' データ最終行=' + lastRowWithData
    );
  });
  Logger.log(log.join('\n'));
}

/**
 * 不良〇月シートを作成(KP・差し戻し統合版)
 * 列構成(A〜W列、ヘッダー1行目・データ2行目〜115行目):
 * タイムスタンプ／処置区分／品証担当者／製造番号／得意先名／品番(図番)／加工者／機種名／設備№／加工数／
 * 良品数(自動計算)／不良数計／不良項目／不良数／不良項目詳細／担当者2／単価／金額(自動計算)／
 * 備考／材質／不良率(自動計算)／キズ原因(非表示・任意)／送信ID(非表示)
 *
 * 【2026-08-10改訂】品証担当者(C列)を先頭グループへ移動。Webアプリ側でログイン中のGoogleアカウントを
 * 自動的に書き込む想定のため、タイムスタンプ・処置区分と同じ「記録の身元情報」としてまとめた
 * (以前はM列にあり、手入力の担当者2と隣接していて紛らわしかった)。
 * 【2026-08-11改訂その1】「数量」→「加工数」に改名(意味が分かりにくいとの指摘)。「良品数」は
 * 加工数−不良数計の自動計算式に変更し、手入力しなくてよいようにした。
 * 【2026-08-11改訂その2】N列「不良項目詳細」を新設(不良項目の自由記述、実測値などを1項目ごとに記録)。
 * 処置区分(B列)は、複数不良項目の2行目以降(追加行)にも同じ値を書き込むよう運用変更
 * (Webアプリ側で対応)。これにより行の色分けが追加行にも及ぶ。あわせて件数集計はB列だけでなく
 * A列(タイムスタンプ、メイン行のみ記入)も条件に含め、追加行を二重カウントしないようにした
 * (下記buildMonthlySummarySheet_参照)。
 * 【2026-08-13改訂】D列「製造番号」を新設(得意先名の左)。QRスキャンで既に取得できていた製造番号を
 * 送信データに含めていながら、これまでシートに保存していなかったため追加した。これにより
 * D列以降(得意先名〜送信ID)が1列ずつ後ろにずれている(既存スプレッドシートは
 * `addManufacturingNumberColumn`で列挿入・データ保持のまま移行する、下記参照)。
 *
 * 【入力ルール】1件の不良につき1行目(メイン行)にA〜J列・L〜O列・P〜T列を入力
 * (処置区分は必須。K列(良品数)は自動計算なので入力不要。単価・金額は社内不良(KP)の場合のみ、
 * 差し戻しは空欄でよい)。同じ不良で不良項目が複数ある場合、2行目以降はB列(処置区分)・
 * M〜P列(不良項目・不良数・詳細)を追加する行を足す。
 */
function buildDefectMonthlySheet_(ss, month) {
  var sheetName = '不良' + month + '月';
  var sheet = replaceSheet_(ss, sheetName);

  var headers = [
    'タイムスタンプ', '処置区分', '品証担当者', '製造番号', '得意先名', '品番(図番)', '加工者', '機種名', '設備№',
    '加工数', '良品数', '不良数計',
    '不良項目', '不良数', '不良項目詳細',
    '担当者2',
    '単価', '金額', '備考', '材質', '不良率',
    'キズ原因(任意)', '送信ID'
  ];
  var lastCol = headers.length;
  sheet.getRange(1, 1, 1, lastCol).setValues([headers])
    .setFontWeight('bold').setBackground(COLOR.HEADER_BG).setFontColor(COLOR.HEADER_FONT);
  sheet.setFrozenRows(1);

  // データ行(2〜115行)に1行おきの縞模様。処置区分で色分けされる行はこの上から上書きされる。
  var bandColors = [];
  for (var r = 0; r < 114; r++) {
    var rowColor = (r % 2 === 0) ? '#FFFFFF' : COLOR.BAND_ALT;
    var row = [];
    for (var c = 0; c < lastCol; c++) row.push(rowColor);
    bandColors.push(row);
  }
  sheet.getRange(2, 1, 114, lastCol).setBackgrounds(bandColors);

  // B列(処置区分)にプルダウン
  var shochiRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SHOCHI_KUBUN, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, 114, 1).setDataValidation(shochiRule);

  // M列(不良項目)にプルダウン
  var itemNames = DEFECT_ITEMS.map(function (item) { return item.name; });
  var itemRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(itemNames, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 13, 114, 1).setDataValidation(itemRule);

  // 処置区分(B列)の値に応じて行全体を色分け(社内不良(KP)=薄青地に紺文字、差し戻し=薄橙地に茶文字)。
  // V・W列(非表示列)はこの色分けの対象外にする(非表示なので見た目上は影響しないが、範囲を広げる意味がないため)。
  var fullRowRange = sheet.getRange(2, 1, 114, lastCol - 2);
  var kpRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B2="社内不良(KP)"')
    .setBackground(COLOR.KP_BG)
    .setFontColor(COLOR.KP_FONT)
    .setRanges([fullRowRange])
    .build();
  var reworkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B2="差し戻し"')
    .setBackground(COLOR.REWORK_BG)
    .setFontColor(COLOR.REWORK_FONT)
    .setRanges([fullRowRange])
    .build();
  sheet.setConditionalFormatRules([kpRule, reworkRule]);

  // 表全体に薄いグリッド罫線、ヘッダー下端は太い罫線で区切る
  sheet.getRange(1, 1, 115, lastCol)
    .setBorder(true, true, true, true, true, true, COLOR.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(1, 1, 1, lastCol)
    .setBorder(null, null, true, null, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // V列(キズ原因)・W列(送信ID)は運用上見る必要がない列のため、列ごと非表示にする
  // (以前は薄いグレー文字で「見えにくくする」だけだったが、行の色分け(背景・文字色)の対象外にした
  // ことで他の列と見た目が不揃いになってしまっていたため、素直に列を隠す方式に変更した)
  sheet.hideColumns(22, 2);

  sheet.setColumnWidth(1, 130); // タイムスタンプ
  sheet.setColumnWidth(2, 120); // 処置区分
  sheet.setColumnWidth(3, 170); // 品証担当者(ログインアカウントで確認した氏名を表示)
  sheet.setColumnWidth(4, 140); // 製造番号
  sheet.setColumnWidth(6, 150); // 品番(図番)
  sheet.setColumnWidth(13, 220); // 不良項目
  sheet.setColumnWidth(15, 240); // 不良項目詳細
  sheet.setColumnWidth(16, 140); // 担当者2(不良項目ごとに原因を作った担当者が違う場合のみ入力する任意項目)

  // K列(良品数)・R列(金額)・U列(不良率)は自動計算式
  var ryohinFormulas = [];
  var amountFormulas = [];
  var rateFormulas = [];
  for (var r = 2; r <= 115; r++) {
    ryohinFormulas.push(['=IFERROR(J' + r + '-L' + r + ',"")']);
    amountFormulas.push(['=IFERROR(L' + r + '*Q' + r + ',"")']);
    rateFormulas.push(['=IFERROR(L' + r + '/J' + r + ',"")']);
  }
  sheet.getRange(2, 11, 114, 1).setFormulas(ryohinFormulas);
  sheet.getRange(2, 18, 114, 1).setFormulas(amountFormulas);
  sheet.getRange(2, 21, 114, 1).setFormulas(rateFormulas).setNumberFormat('0.00%');

  // 見やすさ: シート全体を垂直方向中央揃え、ヘッダー行は水平方向も中央揃え
  sheet.getRange(1, 1, 115, lastCol).setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, lastCol).setHorizontalAlignment('center');
}

/**
 * 不良集計シート(項目別×月別、処置区分を問わず合算)
 * 各月シートの「不良項目」(L列)を件数(COUNTIF)・「不良数」(M列)を個数(SUMIF)で集計。
 * 月ごとに件数・個数の2列(2026-08-12追加)、最下行に月ごとの合計行(2026-08-12追加)。
 * 分類ごとに行の色を分ける。
 */
function buildDefectSummarySheet_(ss) {
  var savedWidths = captureColumnWidths_(ss, '不良集計'); // 手動で調整した列幅があれば覚えておき、作り直した後も引き継ぐ
  var sheet = replaceSheet_(ss, '不良集計');
  buildItemSummarySheet_(sheet, {
    items: DEFECT_ITEMS,
    labelHeaders: ['不良項目', '分類'],
    labelValues: function (item) { return [item.name, item.group]; },
    groupOf: function (item) { return item.group; },
    nameCellCol: 1, // 集計の照合キーは1列目(不良項目名)
    countFormula: function (monthSheetName, nameCell) {
      return '=COUNTIF(\'' + monthSheetName + '\'!M2:M115,' + nameCell + ')';
    },
    qtyFormula: function (monthSheetName, nameCell) {
      return '=SUMIF(\'' + monthSheetName + '\'!M2:M115,' + nameCell + ',\'' + monthSheetName + '\'!N2:N115)';
    },
    savedColumnWidths: savedWidths
  });
}

/**
 * キズ原因の集計シート(現行「キズ集計管理台帳」相当、任意項目のため参考用)
 * 各月シートの「キズ原因」列(V列)を件数(COUNTIF)で、対応する不良数計(L列)を個数(SUMIF)で集計。
 * 月ごとに件数・個数の2列(2026-08-12追加)、最下行に月ごとの合計行(2026-08-12追加)。
 * 原因コード(A〜H)ごとに行の色を分ける。
 */
function buildKpCauseSummarySheet_(ss) {
  var savedWidths = captureColumnWidths_(ss, '不良集計(キズ原因)'); // 手動で調整した列幅があれば覚えておき、作り直した後も引き継ぐ
  var sheet = replaceSheet_(ss, '不良集計(キズ原因)');
  buildItemSummarySheet_(sheet, {
    items: KP_CAUSE_ITEMS,
    labelHeaders: ['原因コード', '分類', '原因'],
    labelValues: function (item) { return [item.code, item.group, item.name]; },
    groupOf: function (item) { return item.code; },
    nameCellCol: 3, // 集計の照合キーは3列目(原因名)
    countFormula: function (monthSheetName, nameCell) {
      return '=COUNTIF(\'' + monthSheetName + '\'!V2:V115,"*"&' + nameCell + ')';
    },
    qtyFormula: function (monthSheetName, nameCell) {
      return '=SUMIF(\'' + monthSheetName + '\'!V2:V115,"*"&' + nameCell + ',\'' + monthSheetName + '\'!L2:L115)';
    },
    savedColumnWidths: savedWidths
  });
}

/**
 * クレーム集計シート(2026-08-17追加)
 * 旧システムの「グラフ」シートB78:U91(客先クレーム件数、月×分類18種)に相当する集計を、
 * ユーザーが旧システム形式で手作りした「客先クレーム管理台帳(CC)」(K列=クレーム内容分類、
 * C列=発生日)から、新マスタ(DEFECT_ITEMS、40項目)ベースで作り直したもの。
 * CC台帳は「不良〇月」のような月別分割シートではなく通し番号の1本シートのため、COUNTIF方式ではなく
 * SUMPRODUCT(MONTH(日付範囲)=対象月)で月ごとの件数を数える。個数・金額は無く件数のみ(旧グラフと同じ)。
 * 不良集計・不良集計(キズ原因)と違って月ごとに1列(件数のみ)なのでbuildItemSummarySheet_は使わない。
 */
function buildClaimSummarySheet_(ss) {
  var savedWidths = captureColumnWidths_(ss, 'クレーム集計');
  var sheet = replaceSheet_(ss, 'クレーム集計');

  var labelHeaders = ['不良項目', '分類'];
  var labelCols = labelHeaders.length;
  var monthStartCol = labelCols + 1;
  var totalCol = monthStartCol + MONTHS.length;
  var lastCol = totalCol;
  var itemStartRow = 2;
  var totalRow = itemStartRow + DEFECT_ITEMS.length;
  var lastRow = totalRow;

  var header = new Array(lastCol).fill('');
  labelHeaders.forEach(function (h, i) { header[i] = h; });
  MONTHS.forEach(function (month, mi) { header[monthStartCol - 1 + mi] = month + '月'; });
  header[totalCol - 1] = '年計';
  sheet.getRange(1, 1, 1, lastCol).setValues([header])
    .setFontWeight('bold').setBackground(COLOR.HEADER_BG).setFontColor(COLOR.HEADER_FONT);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(labelCols);

  var groupColor = buildGroupColorMap_(DEFECT_ITEMS.map(function (item) { return item.group; }));
  var ccSheet = ss.getSheetByName(CC_LEDGER_SHEET_NAME);
  var ccCols = ccSheet ? resolveCcLedgerColumns_(ccSheet) : { dateCol: 2, categoryCol: 3, dataStartRow: 2 };
  var ccRef = "'" + CC_LEDGER_SHEET_NAME + "'!";
  var dateRange = ccRef + '$' + columnToLetter_(ccCols.dateCol) + '$' + ccCols.dataStartRow + ':$' + columnToLetter_(ccCols.dateCol) + '$' + CC_DATA_END_ROW;
  var categoryRange = ccRef + '$' + columnToLetter_(ccCols.categoryCol) + '$' + ccCols.dataStartRow + ':$' + columnToLetter_(ccCols.categoryCol) + '$' + CC_DATA_END_ROW;

  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true, COLOR.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  DEFECT_ITEMS.forEach(function (item, i) {
    var row = itemStartRow + i;
    var nameCell = 'A' + row;
    sheet.getRange(row, 1, 1, labelCols).setValues([[item.name, item.group]]);

    MONTHS.forEach(function (month, mi) {
      var col = monthStartCol + mi;
      // 空欄行(発生日が空)をMONTH()=12月と誤判定しないよう、発生日が入っている行だけを対象にする。
      // また発生日の範囲に「日付として解析できない文字列」が1件でも混じっていると、SUMPRODUCTは
      // 掛け算で0倍される前にMONTH()自体が#VALUE!を返し配列全体に伝播してしまう(実機で確認、
      // 2026-08-17)。IFERRORでそのセルだけ0(どの月とも一致しない値)に読み替えて回避する。
      var formula = '=SUMPRODUCT((' + dateRange + '<>"")*(IFERROR(MONTH(' + dateRange + '),0)=' + month + ')*(' + categoryRange + '=' + nameCell + '))';
      sheet.getRange(row, col).setFormula(formula);
    });

    var monthCells = [];
    for (var mi2 = 0; mi2 < MONTHS.length; mi2++) monthCells.push(columnToLetter_(monthStartCol + mi2) + row);
    sheet.getRange(row, totalCol).setFormula('=' + monthCells.join('+'));

    var groupBase = groupColor[item.group];
    sheet.getRange(row, 1, 1, lastCol).setBackground(groupBase);

    if (i > 0 && item.group !== DEFECT_ITEMS[i - 1].group) {
      sheet.getRange(row, 1, 1, lastCol)
        .setBorder(true, null, null, null, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  });

  sheet.getRange(totalRow, 1).setValue('合計');
  sheet.getRange(totalRow, 1, 1, labelCols).merge();
  for (var col = monthStartCol; col <= totalCol; col++) {
    var colLetter = columnToLetter_(col);
    sheet.getRange(totalRow, col).setFormula('=SUM(' + colLetter + itemStartRow + ':' + colLetter + (totalRow - 1) + ')');
  }
  sheet.getRange(totalRow, 1, 1, lastCol)
    .setBackground(COLOR.SUMMARY_TOTAL_BG).setFontColor(COLOR.TOTAL_FONT).setFontWeight('bold')
    .setBorder(true, null, null, null, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.DOUBLE);

  sheet.getRange(1, labelCols, lastRow, 1)
    .setBorder(null, null, null, true, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(1, totalCol - 1, lastRow, 1)
    .setBorder(null, null, null, true, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.DOUBLE);

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidths(monthStartCol, MONTHS.length + 1, 45);

  if (savedWidths && savedWidths.length === lastCol) {
    savedWidths.forEach(function (w, idx) { sheet.setColumnWidth(idx + 1, w); });
  }

  sheet.getRange(1, 1, lastRow, lastCol).setVerticalAlignment('middle').setHorizontalAlignment('center');
}

/**
 * 【1回だけ手動実行】現行年度の「クレーム集計」シートを追加する(2026-08-17新設)。
 * setupQualityDefectSystemFor_の一部としてbuildClaimSummarySheet_を呼ぶよう組み込み済みのため、
 * 来年度以降の新規スプレッドシートには自動的に含まれる。今年度分だけこれで追加する。
 */
function addClaimSummarySheet() {
  var ss = getCurrentYearSpreadsheet_();
  buildClaimSummarySheet_(ss);
  SpreadsheetApp.flush();
  Logger.log('「クレーム集計」シートを追加しました。');
}

/**
 * 【1回だけ手動実行】「客先クレーム管理台帳(CC)」のクレーム内容分類列に、DEFECT_ITEMSの
 * 項目名からなるプルダウン(データの入力規則)を追加する(2026-08-17新設)。
 * 【2026-08-18】列位置はresolveCcLedgerColumns_で動的に求めるようにしたため、buildLedgerSheetV2_
 * 導入後の新デザインに切り替わっても対応できる(新デザインは元々このプルダウンを内蔵しているため
 * 実質不要になるが、旧デザインのまま運用する場合の保険として残す)。
 * setAllowInvalid(true)にしてあるため、既存の入力値やマスタに無い値を書いても壊れない
 * (プルダウンの候補として出るだけで、直接入力も引き続きできる)。
 */
function addCcClaimCategoryValidation() {
  var ss = getCurrentYearSpreadsheet_();
  var sheet = ss.getSheetByName(CC_LEDGER_SHEET_NAME);
  if (!sheet) throw new Error('「' + CC_LEDGER_SHEET_NAME + '」シートが見つかりません');

  var ccCols = resolveCcLedgerColumns_(sheet);
  if (ccCols.categoryCol === -1) throw new Error('「クレーム内容分類」列が見つかりません');

  var itemNames = DEFECT_ITEMS.map(function (item) { return item.name; });
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(itemNames, true)
    .setAllowInvalid(true)
    .build();
  var rows = 150; // 台帳の伸びしろとして150行分
  sheet.getRange(ccCols.dataStartRow, ccCols.categoryCol, rows, 1).setDataValidation(rule);
  Logger.log('「' + CC_LEDGER_SHEET_NAME + '」' + columnToLetter_(ccCols.categoryCol) + '列(' +
    ccCols.dataStartRow + '〜' + (ccCols.dataStartRow + rows - 1) + '行目)にプルダウンを追加しました。');
}

/**
 * 【1回だけ手動実行】「社内不良管理台帳(製造工程)(SK)」のC列(発生日(品質会議開催日))に、
 * CC/KP台帳と同じ日付の入力規則(カレンダーアイコンから選択可能)・表示形式を設定する(2026-08-17新設)。
 * debugInspectDateColumnValidation(MigrateOldData.gs)の実機調査で、この列だけ入力規則が無く
 * 表示形式も日付になっていない(0.###############、生の数値書式)ことが判明したための対応。
 * 既存の入力済みデータ(値そのもの)は一切変更しない(表示形式・入力規則のみ変更)。
 * setAllowInvalid(true)にしてあるため、既存データが仮に日付以外の値でもエラー扱いにはならない。
 */
function fixSkDateColumnValidation() {
  var ss = getCurrentYearSpreadsheet_();
  var sheet = ss.getSheetByName(SK_LEDGER_SHEET_NAME);
  if (!sheet) throw new Error('「' + SK_LEDGER_SHEET_NAME + '」シートが見つかりません');

  var rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('日付を入力してください(例: 2026/8/17)。セルを選択すると右端に出るカレンダーアイコンからも選べます。')
    .build();
  var rows = 150; // 台帳の伸びしろ(buildImprovementLedgerSheet_のdataRowsと同じ考え方)
  var range = sheet.getRange(SK_DATA_START_ROW, SK_DATE_COL, rows, 1);
  range.setDataValidation(rule);
  range.setNumberFormat('yyyy"年"m"月"d"日"'); // CC/KPの発生日列と同じ表示形式に揃える

  Logger.log('「' + SK_LEDGER_SHEET_NAME + '」C列(' + SK_DATA_START_ROW + '〜' + (SK_DATA_START_ROW + rows - 1) + '行目)に日付検証・表示形式を設定しました。');
}

/**
 * シートが存在すれば、その時点の列幅(1列目〜最終列)を配列で返す。存在しなければnull。
 * replaceSheet_で作り直す前に呼び、手動で調整した列幅を後で復元できるようにする。
 */
function captureColumnWidths_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return null;
  var widths = [];
  for (var c = 1; c <= lastCol; c++) widths.push(sheet.getColumnWidth(c));
  return widths;
}

/**
 * 不良集計・不良集計(キズ原因)で共通の「項目×月」集計シートを組み立てる。
 * 列構成: ラベル列(可変数) → 月ごとに[件数,個数]の2列×12ヶ月 → 年計[件数,個数]の2列。
 * 行構成: ヘッダー2行(月をまたぐ結合ヘッダー) → 項目の数だけデータ行 → 最下行に月ごとの合計行。
 */
function buildItemSummarySheet_(sheet, cfg) {
  var labelCols = cfg.labelHeaders.length;
  var monthStartCol = labelCols + 1;
  var totalStartCol = monthStartCol + MONTHS.length * 2;
  var lastCol = totalStartCol + 1;
  var itemStartRow = 3;
  var totalRow = itemStartRow + cfg.items.length;
  var lastRow = totalRow;

  // ヘッダー行1: ラベル列見出し、月ごとの「〇月」、「年計」(まだマージしない)
  var header1 = new Array(lastCol).fill('');
  cfg.labelHeaders.forEach(function (h, i) { header1[i] = h; });
  MONTHS.forEach(function (month, mi) { header1[monthStartCol - 1 + mi * 2] = month + '月'; });
  header1[totalStartCol - 1] = '年計';
  sheet.getRange(1, 1, 1, lastCol).setValues([header1]);

  // ヘッダー行2: 月・年計それぞれの下に「件数」「個数」(ラベル列は空欄のまま)
  var header2 = new Array(lastCol).fill('');
  for (var mi3 = 0; mi3 < MONTHS.length; mi3++) {
    header2[monthStartCol - 1 + mi3 * 2] = '件数';
    header2[monthStartCol - 1 + mi3 * 2 + 1] = '個数';
  }
  header2[totalStartCol - 1] = '件数';
  header2[totalStartCol] = '個数';
  sheet.getRange(2, 1, 1, lastCol).setValues([header2]);

  // 両方の行に値を書き終えてからマージする(先にマージすると2行目への書き込みで
  // 「結合済みセルを編集しています」エラーになるため)
  for (var lc = 1; lc <= labelCols; lc++) sheet.getRange(1, lc, 2, 1).merge();
  for (var mi2 = 0; mi2 < MONTHS.length; mi2++) sheet.getRange(1, monthStartCol + mi2 * 2, 1, 2).merge();
  sheet.getRange(1, totalStartCol, 1, 2).merge();

  // ヘッダー行1(月名・年計)は濃紺。行2(件数/個数)は同系色(青系)の濃淡で分け、
  // 色相を変えすぎずにどちらの列か分かるようにする(2026-08-12)。
  sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground(COLOR.HEADER_BG).setFontColor(COLOR.HEADER_FONT);
  sheet.getRange(2, 1, 1, lastCol).setFontWeight('bold').setFontColor(COLOR.TOTAL_FONT);
  for (var mi5 = 0; mi5 <= MONTHS.length; mi5++) {
    var blockStart = mi5 < MONTHS.length ? monthStartCol + mi5 * 2 : totalStartCol;
    sheet.getRange(2, blockStart).setBackground(COLOR.SUMMARY_COUNT_BG); // 件数(薄い方)
    sheet.getRange(2, blockStart + 1).setBackground(COLOR.SUMMARY_QTY_BG); // 個数(濃い方)
  }
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(labelCols); // ラベル列まで固定(〇月列の手前)

  var groupColor = buildGroupColorMap_(cfg.items.map(cfg.groupOf));
  var nameColLetter = columnToLetter_(cfg.nameCellCol);

  // 先に表全体へ薄いグリッド罫線を引いておく(この後のグループ区切り線・ヘッダー下線・合計行罫線で上書きする)
  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true, COLOR.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  cfg.items.forEach(function (item, i) {
    var row = itemStartRow + i;
    var nameCell = nameColLetter + row;
    sheet.getRange(row, 1, 1, labelCols).setValues([cfg.labelValues(item)]);

    MONTHS.forEach(function (month, mi) {
      var monthSheetName = '不良' + month + '月';
      var countCol = monthStartCol + mi * 2;
      sheet.getRange(row, countCol).setFormula(cfg.countFormula(monthSheetName, nameCell));
      sheet.getRange(row, countCol + 1).setFormula(cfg.qtyFormula(monthSheetName, nameCell));
    });

    // 年計(件数・個数とも、月ごとの件数列/個数列だけを飛び飛びに合算する式を組み立てる)
    var countTerms = [], qtyTerms = [];
    for (var mi4 = 0; mi4 < MONTHS.length; mi4++) {
      var c = monthStartCol + mi4 * 2;
      countTerms.push(columnToLetter_(c) + row);
      qtyTerms.push(columnToLetter_(c + 1) + row);
    }
    sheet.getRange(row, totalStartCol).setFormula('=' + countTerms.join('+'));
    sheet.getRange(row, totalStartCol + 1).setFormula('=' + qtyTerms.join('+'));

    var group = cfg.groupOf(item);
    var groupBase = groupColor[group]; // GROUP_PALETTEの薄い色をそのグループの基準色として使う
    var countColor = groupBase;                  // 件数=グループの色そのまま(薄い)
    var qtyColor = darkenColor_(groupBase, 7);    // 個数=同じ色相・彩度のまま少し濃くしたもの

    // ラベル列(項目名・分類)はグループの色一色
    sheet.getRange(row, 1, 1, labelCols).setBackground(groupBase);
    // 件数・個数の列も、月ごと・年計ともグループの色の濃淡で塗る(2026-08-12、グループごとに色相を変える)
    for (var mi7 = 0; mi7 <= MONTHS.length; mi7++) {
      var blockStart2 = mi7 < MONTHS.length ? monthStartCol + mi7 * 2 : totalStartCol;
      sheet.getRange(row, blockStart2).setBackground(countColor);
      sheet.getRange(row, blockStart2 + 1).setBackground(qtyColor);
    }

    // グループが切り替わる行の上端だけ太い罫線を引いて、区切りを分かりやすくする(表全体の幅で)
    if (i > 0 && group !== cfg.groupOf(cfg.items[i - 1])) {
      sheet.getRange(row, 1, 1, lastCol)
        .setBorder(true, null, null, null, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  });

  // 最下行: 月ごと・年計それぞれの件数列/個数列を、その列の項目行すべてで合計する
  sheet.getRange(totalRow, 1).setValue('合計');
  if (labelCols > 1) sheet.getRange(totalRow, 1, 1, labelCols).merge();
  for (var col = monthStartCol; col <= totalStartCol + 1; col++) {
    var colLetter = columnToLetter_(col);
    sheet.getRange(totalRow, col).setFormula(
      '=SUM(' + colLetter + itemStartRow + ':' + colLetter + (totalRow - 1) + ')'
    );
  }
  sheet.getRange(totalRow, 1, 1, lastCol)
    .setBackground(COLOR.SUMMARY_TOTAL_BG).setFontColor(COLOR.TOTAL_FONT).setFontWeight('bold')
    // 合計行の上端も、年計と月の境目と同じく二重線にして「ここから先は合計」と分かりやすくする(2026-08-12)
    .setBorder(true, null, null, null, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.DOUBLE);

  sheet.getRange(1, 1, 2, lastCol)
    .setBorder(null, null, true, null, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 月ブロックの境目に縦線を引いて、どこからどこまでが同じ月か分かりやすくする(2026-08-12)。
  // ラベル列とデータ列の境目も同じ扱いにする。
  sheet.getRange(1, labelCols, lastRow, 1)
    .setBorder(null, null, null, true, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  for (var mi6 = 0; mi6 < MONTHS.length; mi6++) {
    var boundaryCol = monthStartCol + mi6 * 2 + 1; // その月ブロックの右端の列
    var isLastMonth = mi6 === MONTHS.length - 1; // 年計との境目だけは二重線で強調する(2026-08-12)
    sheet.getRange(1, boundaryCol, lastRow, 1)
      .setBorder(null, null, null, true, null, null,
        isLastMonth ? COLOR.HEADER_BORDER : COLOR.GRID_BORDER,
        isLastMonth ? SpreadsheetApp.BorderStyle.DOUBLE : SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  sheet.setColumnWidth(cfg.nameCellCol, 220);
  sheet.setColumnWidths(monthStartCol, MONTHS.length * 2 + 2, 45); // 件数・個数・年計列を狭く揃える(既定値)

  // 手動で調整した列幅が記録されていれば、既定値の代わりにそれを復元する
  // (列数が変わっていない場合のみ。項目数やヘッダー構成が変わって列数がずれた場合は既定値のまま)
  if (cfg.savedColumnWidths && cfg.savedColumnWidths.length === lastCol) {
    cfg.savedColumnWidths.forEach(function (w, idx) { sheet.setColumnWidth(idx + 1, w); });
  }

  // 見やすさ: シート全体を垂直・水平方向とも中央揃え
  sheet.getRange(1, 1, lastRow, lastCol).setVerticalAlignment('middle').setHorizontalAlignment('center');
}

/** 配列から重複を除き、初出順を保ったまま返す(グラフのグループ一覧作成用) */
function uniqueInOrder_(list) {
  var seen = [];
  list.forEach(function (v) { if (seen.indexOf(v) === -1) seen.push(v); });
  return seen;
}

/** 登場順に GROUP_PALETTE の色を割り当てたマップ({グループ名: 色}) を作る */
function buildGroupColorMap_(groupList) {
  var map = {};
  var next = 0;
  groupList.forEach(function (g) {
    if (!map[g]) {
      map[g] = GROUP_PALETTE[next % GROUP_PALETTE.length];
      next++;
    }
  });
  return map;
}

/**
 * 指定したhex色を、色相・彩度は変えずに明度だけ下げて濃くした色を返す(HSLの明度を下げる)。
 * 黒に混ぜて暗くする方式だと、GROUP_PALETTEのような薄いパステルカラーが色味の抜けた
 * グレーっぽい色になってしまう(色相・彩度の差が小さいまま明るさだけ下がるため)ので、
 * 不良集計等で「グループの色そのもの(件数)」から「少し濃い同じ色味(個数)」を作るのに使う。
 * @param {string} hex 元の色(例: '#E8F0FE')
 * @param {number} amount 明度を下げるポイント数(0〜100、HSLのL値ベース)
 */
function darkenColor_(hex, amount) {
  var num = parseInt(hex.replace('#', ''), 16);
  var r = ((num >> 16) & 0xFF) / 255, g = ((num >> 8) & 0xFF) / 255, b = (num & 0xFF) / 255;
  var max = Math.max(r, g, b), min = Math.min(r, g, b);
  var h = 0, s = 0, l = (max + min) / 2;
  if (max !== min) {
    var d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  l = Math.max(0, l - amount / 100);

  function hue2rgb(p, q, t) {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  }
  var r2, g2, b2;
  if (s === 0) {
    r2 = g2 = b2 = l;
  } else {
    var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    var p = 2 * l - q;
    r2 = hue2rgb(p, q, h + 1 / 3);
    g2 = hue2rgb(p, q, h);
    b2 = hue2rgb(p, q, h - 1 / 3);
  }
  function toHex(x) {
    var v = Math.round(x * 255);
    return ('0' + v.toString(16)).slice(-2);
  }
  return '#' + toHex(r2) + toHex(g2) + toHex(b2);
}

/**
 * 月次サマリーシート
 * 処置区分(B列)ごとに「件数・個数(・金額)」を月別に集計し、両方を合算した月別合計もあわせて表示する。
 * 現行「キズ集計管理台帳」の「月別不良グラフ」相当。KP行・差し戻し行・合計行をそれぞれ色分けする。
 *
 * 【2026-08-11改訂】B列(処置区分)は複数不良項目の追加行にも同じ値が入るようになったため、
 * 件数はB列だけでなくA列(タイムスタンプ、メイン行のみ記入)も条件に加えて、追加行を
 * 二重カウントしないようにしている。個数・金額はメイン行のL列(不良数計)・R列(金額)の合計
 * (追加行はL・R列とも空欄のため、SUMIFSの条件をB列だけにしても二重加算にはならない)。
 */
function buildMonthlySummarySheet_(ss) {
  var sheet = replaceSheet_(ss, '月次サマリー');

  var header = ['項目'].concat(MONTHS.map(function (m) { return m + '月'; })).concat(['年計']);
  var lastMonthCol = 1 + MONTHS.length;
  var totalCol = lastMonthCol + 1;
  sheet.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground(COLOR.HEADER_BG).setFontColor(COLOR.HEADER_FONT);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  var kpCountRow = 2, kpQtyRow = 3, kpAmountRow = 4;
  var reworkCountRow = 6, reworkQtyRow = 7;
  var totalCountRow = 9, totalQtyRow = 10;

  var labels = {};
  labels[kpCountRow] = 'KP 不良件数';
  labels[kpQtyRow] = 'KP 不良個数';
  labels[kpAmountRow] = 'KP 不良金額';
  labels[reworkCountRow] = '差し戻し 件数';
  labels[reworkQtyRow] = '差し戻し 個数';
  labels[totalCountRow] = '合計 不良件数(KP＋差し戻し)';
  labels[totalQtyRow] = '合計 不良個数(KP＋差し戻し)';

  var rowColor = {};
  var rowFont = {};
  [kpCountRow, kpQtyRow, kpAmountRow].forEach(function (r) { rowColor[r] = COLOR.SUMMARY_KP_BG; rowFont[r] = COLOR.KP_FONT; });
  [reworkCountRow, reworkQtyRow].forEach(function (r) { rowColor[r] = COLOR.SUMMARY_REWORK_BG; rowFont[r] = COLOR.REWORK_FONT; });
  [totalCountRow, totalQtyRow].forEach(function (r) { rowColor[r] = COLOR.SUMMARY_TOTAL_BG; rowFont[r] = COLOR.TOTAL_FONT; });

  // 先に表全体へ薄いグリッド罫線を引いておく(この後のブロック区切り線・ヘッダー下線で上書きする)
  var lastRow = totalQtyRow;
  sheet.getRange(1, 1, lastRow, totalCol)
    .setBorder(true, true, true, true, true, true, COLOR.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  Object.keys(labels).forEach(function (r) {
    var row = Number(r);
    sheet.getRange(row, 1).setValue(labels[r]).setFontWeight('bold');
    sheet.getRange(row, 1, 1, totalCol).setBackground(rowColor[row]).setFontColor(rowFont[row]);
  });

  // KP・差し戻し・合計、各ブロックの先頭行の上端に太い罫線を引いて区切りを分かりやすくする
  [kpCountRow, reworkCountRow, totalCountRow].forEach(function (row) {
    sheet.getRange(row, 1, 1, totalCol)
      .setBorder(true, null, null, null, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.getRange(1, 1, 1, totalCol)
    .setBorder(null, null, true, null, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  MONTHS.forEach(function (month, mi) {
    var col = 2 + mi;
    var colLetter = columnToLetter_(col);
    var sn = '\'不良' + month + '月\'';

    sheet.getRange(kpCountRow, col).setFormula('=COUNTIFS(' + sn + '!A2:A115,"<>",' + sn + '!B2:B115,"社内不良(KP)")');
    sheet.getRange(kpQtyRow, col).setFormula('=SUMIFS(' + sn + '!L2:L115,' + sn + '!B2:B115,"社内不良(KP)")');
    sheet.getRange(kpAmountRow, col).setFormula('=SUMIFS(' + sn + '!R2:R115,' + sn + '!B2:B115,"社内不良(KP)")');
    sheet.getRange(reworkCountRow, col).setFormula('=COUNTIFS(' + sn + '!A2:A115,"<>",' + sn + '!B2:B115,"差し戻し")');
    sheet.getRange(reworkQtyRow, col).setFormula('=SUMIFS(' + sn + '!L2:L115,' + sn + '!B2:B115,"差し戻し")');
    sheet.getRange(totalCountRow, col).setFormula('=' + colLetter + kpCountRow + '+' + colLetter + reworkCountRow);
    sheet.getRange(totalQtyRow, col).setFormula('=' + colLetter + kpQtyRow + '+' + colLetter + reworkQtyRow);
  });

  Object.keys(labels).forEach(function (r) {
    var row = Number(r);
    sheet.getRange(row, totalCol).setFormula(
      '=SUM(B' + row + ':' + columnToLetter_(lastMonthCol) + row + ')'
    );
  });

  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidths(2, MONTHS.length, 50); // 〇月列を狭く

  // 見やすさ: シート全体を垂直方向中央揃え、ヘッダー行と見出し列(項目名)は水平方向も中央揃え
  sheet.getRange(1, 1, lastRow, totalCol).setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, totalCol).setHorizontalAlignment('center');
  sheet.getRange(1, 1, lastRow, 1).setHorizontalAlignment('center');
}

/**
 * グラフシート(ダッシュボード)。
 * 旧「客先クレーム(CC)・社内不良(KP)管理台帳」の「グラフ」シート(棒・折れ線8枚、うち4枚が
 * 「社内」と「出荷検査」で内容が重複していて分かりにくかった)を、新システムの集計データ
 * (不良集計・不良集計(キズ原因)・月次サマリー)から作り直したもの。新システムには「出荷検査」
 * という区分が無い(社内不良(KP)/差し戻しのみ)ため、その区別は再現していない。
 * 2026-08-12、ユーザー合意の上で以下4種類に整理:
 *   ①月別不良個数(棒グラフ) ②月別不良金額(折れ線、単価を追う社内不良(KP)のみ)
 *   ③分類別内訳(積み上げ棒、月別×5グループ) ④キズ原因内訳(棒グラフ、原因グループ別の年計個数、任意項目)
 * A1〜O.. に各グラフの元データ(数式による自動集計、編集不要)を置き、その下にグラフ本体を配置する。
 * シートはタブの先頭(1番目)に配置する。
 */
function buildDefectDashboardSheet_(ss) {
  var sheet = replaceSheet_(ss, 'グラフ');

  sheet.getRange(1, 1, 1, 15).merge().setValue('■ グラフ元データ(自動計算・編集不要)')
    .setFontWeight('bold').setBackground(COLOR.HEADER_BG).setFontColor(COLOR.HEADER_FONT);

  var monthlySummarySheetName = "'月次サマリー'";
  var itemSummarySheetName = "'不良集計'";
  var causeSummarySheetName = "'不良集計(キズ原因)'";

  // --- 表1: 月別不良個数(A3:B15、月次サマリーの「合計 不良個数(KP＋差し戻し)」行を参照) ---
  sheet.getRange(3, 1, 1, 2).setValues([['月', '不良個数']]).setFontWeight('bold');
  MONTHS.forEach(function (month, mi) {
    var row = 4 + mi;
    var colLetter = columnToLetter_(2 + mi); // 月次サマリーの月別列(B〜M)
    sheet.getRange(row, 1).setFormula('=' + monthlySummarySheetName + '!' + colLetter + '1');
    sheet.getRange(row, 2).setFormula('=' + monthlySummarySheetName + '!' + colLetter + '10');
  });

  // --- 表2: 月別不良金額(D3:E15、月次サマリーの「KP 不良金額」行を参照。差し戻しは単価を追わないため対象外) ---
  sheet.getRange(3, 4, 1, 2).setValues([['月', '不良金額(社内不良KP)']]).setFontWeight('bold');
  MONTHS.forEach(function (month, mi) {
    var row = 4 + mi;
    var colLetter = columnToLetter_(2 + mi);
    sheet.getRange(row, 4).setFormula('=' + monthlySummarySheetName + '!' + colLetter + '1');
    sheet.getRange(row, 5).setFormula('=' + monthlySummarySheetName + '!' + colLetter + '4');
  });

  // --- 表3: 分類別内訳(月別×5グループ、G3:L15、不良集計シートの項目別個数をSUMIFで分類ごとに合算) ---
  var defectGroups = uniqueInOrder_(DEFECT_ITEMS.map(function (item) { return item.group; }));
  var defectGroupColor = buildGroupColorMap_(DEFECT_ITEMS.map(function (item) { return item.group; }));
  var table3Header = ['月'].concat(defectGroups);
  sheet.getRange(3, 7, 1, table3Header.length).setValues([table3Header]).setFontWeight('bold');
  MONTHS.forEach(function (month, mi) {
    var row = 4 + mi;
    var qtyColLetter = columnToLetter_(4 + mi * 2); // 不良集計シートの月別「個数」列(D,F,H...)
    sheet.getRange(row, 7).setFormula('=' + monthlySummarySheetName + '!' + columnToLetter_(2 + mi) + '1');
    defectGroups.forEach(function (group, gi) {
      sheet.getRange(row, 8 + gi).setFormula(
        '=SUMIF(' + itemSummarySheetName + '!$B$3:$B$' + (2 + DEFECT_ITEMS.length) + ',"' + group + '",'
        + itemSummarySheetName + '!' + qtyColLetter + '$3:' + qtyColLetter + '$' + (2 + DEFECT_ITEMS.length) + ')'
      );
    });
  });

  // --- 表4: キズ原因内訳(原因グループ別の年計個数、N3:O11、不良集計(キズ原因)シートをSUMIFで原因グループごとに合算) ---
  var causeGroups = uniqueInOrder_(KP_CAUSE_ITEMS.map(function (item) { return item.group; }));
  var causeLastRow = 2 + KP_CAUSE_ITEMS.length;
  var causeYearQtyCol = columnToLetter_(4 + MONTHS.length * 2 + 1); // 不良集計(キズ原因)シートの「年計」個数列
  sheet.getRange(3, 14, 1, 2).setValues([['原因グループ', '年計個数']]).setFontWeight('bold');
  causeGroups.forEach(function (group, gi) {
    var row = 4 + gi;
    sheet.getRange(row, 14).setValue(group);
    sheet.getRange(row, 15).setFormula(
      '=SUMIF(' + causeSummarySheetName + '!$B$3:$B$' + causeLastRow + ',"' + group + '",'
      + causeSummarySheetName + '!$' + causeYearQtyCol + '$3:$' + causeYearQtyCol + '$' + causeLastRow + ')'
    );
  });

  // --- グラフ本体 ---
  var chart1 = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(3, 1, 1 + MONTHS.length, 2))
    .setPosition(20, 1, 0, 0)
    .setOption('title', '月別不良個数')
    .setOption('width', 480).setOption('height', 300)
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COLOR.HEADER_BG])
    .build();
  sheet.insertChart(chart1);

  var chart2 = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(3, 4, 1 + MONTHS.length, 2))
    .setPosition(20, 9, 0, 0)
    .setOption('title', '月別不良金額(社内不良(KP)のみ)')
    .setOption('width', 480).setOption('height', 300)
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COLOR.REWORK_FONT])
    .build();
  sheet.insertChart(chart2);

  var chart3 = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(3, 7, 1 + MONTHS.length, 1 + defectGroups.length))
    .setPosition(40, 1, 0, 0)
    .setOption('title', '分類別内訳(月別・積み上げ)')
    .setOption('width', 480).setOption('height', 300)
    .setOption('isStacked', true)
    .setOption('colors', defectGroups.map(function (g) { return defectGroupColor[g]; }))
    .build();
  sheet.insertChart(chart3);

  var chart4 = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(3, 14, 1 + causeGroups.length, 2))
    .setPosition(40, 9, 0, 0)
    .setOption('title', 'キズ原因内訳(年計個数、任意項目)')
    .setOption('width', 480).setOption('height', 300)
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COLOR.KP_FONT])
    .build();
  sheet.insertChart(chart4);

  sheet.setColumnWidths(1, 15, 90);
  sheet.getRange(1, 1, 60, 15).setVerticalAlignment('middle');

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1); // タブの先頭に配置

  return sheet;
}

/**
 * 【手動実行用】「グラフ」シートを新規追加する(2026-08-12)。既存の「不良〇月」12シートの
 * 入力済みデータには一切影響しない(グラフシートを作り直すだけ)。
 */
function addDefectDashboardSheet() {
  var ss = getCurrentYearSpreadsheet_();
  buildDefectDashboardSheet_(ss);
  SpreadsheetApp.flush();
  Logger.log('「グラフ」シートを作成し、先頭タブに配置しました。');
}

/**
 * 改善計画書台帳(CC/KP/SK)共通の土台(2026-08-18、旧buildImprovementLedgerSheet_を全面作り直し)。
 * 記録情報欄(recordFields)＋LEDGER_WORKFLOW_STEPSの9ステップ(担当+完了日)＋水平展開の判定＋
 * 差し戻し文書No.＋備考、で1シートを組み立てる。台帳番号は手入力(自動採番はしない、従来通り)。
 * 【旧デザインからの変更点】①完了日・発生日は単一セル+日付検証(旧:⇒矢印+月/日プルダウンの3列、
 * 年の記録なし)②担当は全ステップでプルダウン化(旧:一部のみ)③ステップごとに色分け
 * ④2行ヘッダー(グループ行+列見出し行)で見やすく。
 *
 * @param {string} sheetName シート名
 * @param {Array<{header:string, dropdown?:string[]}>} recordFields 記録情報欄の列定義(先頭から順)
 * @param {string[]} staffList ワークフロー「担当」列(全9ステップ共通)のプルダウン選択肢
 * @param {Object} [formulaCols] 金額の自動計算をする場合、{qty:'数量列の見出し名', unitPrice:'単価列の見出し名', amount:'金額列の見出し名'}
 */
function buildLedgerSheetV2_(ss, sheetName, recordFields, staffList, formulaCols) {
  var savedWidths = captureColumnWidths_(ss, sheetName);
  var sheet = replaceSheet_(ss, sheetName);

  var recordHeaders = recordFields.map(function (f) { return f.header; });
  var recordColCount = recordHeaders.length;
  var stepStartCol = recordColCount + 1;
  var tailHeaders = ['水平展開の判定(〇/×)', '差し戻し文書No.(任意)', '備考'];
  var tailStartCol = stepStartCol + LEDGER_WORKFLOW_STEPS.length * 2;
  var lastCol = tailStartCol + tailHeaders.length - 1;
  var dataStartRow = 3; // 1行目=グループ見出し、2行目=列見出し、3行目からデータ
  var dataRows = 150;   // 通し番号の台帳のため月別分割なし。150行を超える場合は手動で行を追加する
  var lastDataRow = dataStartRow + dataRows - 1;
  var stepNumerals = '①②③④⑤⑥⑦⑧⑨';
  var stepPalette = GROUP_PALETTE.concat(['#FDE2E2']); // GROUP_PALETTE(8色)+1色で9ステップ分を用意

  // --- ヘッダー行1(グループ見出し)・行2(列見出し) ---
  var group1 = new Array(lastCol).fill('');
  group1[0] = '記録情報';
  LEDGER_WORKFLOW_STEPS.forEach(function (step, i) {
    group1[stepStartCol - 1 + i * 2] = stepNumerals.charAt(i) + ' ' + step;
  });
  group1[tailStartCol - 1] = '判定・備考';

  var group2 = recordHeaders.slice();
  LEDGER_WORKFLOW_STEPS.forEach(function () { group2.push('担当', '完了日'); });
  group2 = group2.concat(tailHeaders);

  sheet.getRange(1, 1, 1, lastCol).setValues([group1]);
  sheet.getRange(2, 1, 1, lastCol).setValues([group2]);

  // 両方の行に値を書き終えてからマージする(先にマージすると2行目への書き込みでエラーになるため)
  sheet.getRange(1, 1, 1, recordColCount).merge();
  LEDGER_WORKFLOW_STEPS.forEach(function (step, i) { sheet.getRange(1, stepStartCol + i * 2, 1, 2).merge(); });
  sheet.getRange(1, tailStartCol, 1, tailHeaders.length).merge();

  sheet.getRange(1, 1, 2, lastCol).setFontWeight('bold').setBackground(COLOR.HEADER_BG).setFontColor(COLOR.HEADER_FONT);
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(3); // 台帳番号・発生日・関連文書あたりまで固定してワークフロー欄をスクロールしやすくする

  // --- データ行の縞模様(下地。ステップ欄は後でこの上から塗りつぶす) ---
  var bandColors = [];
  for (var r = 0; r < dataRows; r++) {
    var rowColor = (r % 2 === 0) ? '#FFFFFF' : COLOR.BAND_ALT;
    var rowArr = [];
    for (var c = 0; c < lastCol; c++) rowArr.push(rowColor);
    bandColors.push(rowArr);
  }
  sheet.getRange(dataStartRow, 1, dataRows, lastCol).setBackgrounds(bandColors);

  // --- ステップごとの色分け(ヘッダー・データ行とも、ステップ全体を単色で塗る) ---
  LEDGER_WORKFLOW_STEPS.forEach(function (step, i) {
    var col = stepStartCol + i * 2;
    var color = stepPalette[i % stepPalette.length];
    sheet.getRange(1, col, 2, 2).setBackground(color).setFontColor(COLOR.TOTAL_FONT);
    sheet.getRange(dataStartRow, col, dataRows, 2).setBackground(color);
  });

  // --- 記録情報欄: プルダウン・日付列・列幅 ---
  recordFields.forEach(function (f, i) {
    var col = i + 1;
    if (f.dropdown && f.dropdown.length > 0) {
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(f.dropdown, true).setAllowInvalid(true).build();
      sheet.getRange(dataStartRow, col, dataRows, 1).setDataValidation(rule);
    }
    if (/発生日|開催日/.test(f.header)) {
      applyLedgerDateColumn_(sheet, col, dataStartRow, dataRows);
      sheet.setColumnWidth(col, 110);
    } else if (/詳細|備考/.test(f.header)) {
      sheet.setColumnWidth(col, 220);
    } else if (/関連文書/.test(f.header)) {
      sheet.setColumnWidth(col, 150);
    } else {
      sheet.setColumnWidth(col, 100);
    }
  });

  // --- ワークフロー欄: 担当(全ステップ共通プルダウン)+完了日(単一セル・日付検証) ---
  var staffRule = SpreadsheetApp.newDataValidation().requireValueInList(staffList, true).setAllowInvalid(true).build();
  LEDGER_WORKFLOW_STEPS.forEach(function (step, i) {
    var workerCol = stepStartCol + i * 2;
    var dateCol = workerCol + 1;
    sheet.getRange(dataStartRow, workerCol, dataRows, 1).setDataValidation(staffRule);
    applyLedgerDateColumn_(sheet, dateCol, dataStartRow, dataRows);
    sheet.setColumnWidth(workerCol, 90);
    sheet.setColumnWidth(dateCol, 100);
  });

  // --- 判定(〇/×)・差し戻し文書No.・備考 ---
  var hankoRule = SpreadsheetApp.newDataValidation().requireValueInList(HANKO_CHOICES_, true).setAllowInvalid(true).build();
  sheet.getRange(dataStartRow, tailStartCol, dataRows, 1).setDataValidation(hankoRule);
  sheet.setColumnWidth(tailStartCol, 110);
  sheet.setColumnWidth(tailStartCol + 1, 150);
  sheet.setColumnWidth(tailStartCol + 2, 220);

  // --- 金額の自動計算(数量×単価) ---
  if (formulaCols) {
    var qtyCol = recordHeaders.indexOf(formulaCols.qty) + 1;
    var priceCol = recordHeaders.indexOf(formulaCols.unitPrice) + 1;
    var amountCol = recordHeaders.indexOf(formulaCols.amount) + 1;
    if (qtyCol > 0 && priceCol > 0 && amountCol > 0) {
      var qtyLetter = columnToLetter_(qtyCol), priceLetter = columnToLetter_(priceCol);
      var formulas = [];
      for (var r2 = dataStartRow; r2 <= lastDataRow; r2++) {
        formulas.push(['=IFERROR(' + qtyLetter + r2 + '*' + priceLetter + r2 + ',"")']);
      }
      sheet.getRange(dataStartRow, amountCol, dataRows, 1).setFormulas(formulas);
    }
  }

  // --- 罫線: 表全体に薄いグリッド、ヘッダー下端・記録情報とワークフローの境目は太線 ---
  sheet.getRange(1, 1, lastDataRow, lastCol)
    .setBorder(true, true, true, true, true, true, COLOR.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(1, 1, 2, lastCol)
    .setBorder(null, null, true, null, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(1, recordColCount, lastDataRow, 1)
    .setBorder(null, null, null, true, null, null, COLOR.HEADER_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  LEDGER_WORKFLOW_STEPS.forEach(function (step, i) {
    var boundaryCol = stepStartCol + i * 2 + 1;
    sheet.getRange(1, boundaryCol, lastDataRow, 1)
      .setBorder(null, null, null, true, null, null, COLOR.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  sheet.getRange(1, 1, lastDataRow, lastCol).setVerticalAlignment('middle').setHorizontalAlignment('center');

  if (savedWidths && savedWidths.length === lastCol) {
    savedWidths.forEach(function (w, idx) { sheet.setColumnWidth(idx + 1, w); });
  }

  return sheet;
}

/** 日付列に単一セルの日付検証・表示形式・入力ヒントを設定する(fixSkDateColumnValidationと同じ考え方、2026-08-18) */
function applyLedgerDateColumn_(sheet, col, startRow, rows) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('日付を入力してください(例: 2026/8/17)。セルを選択すると右端に出るカレンダーアイコンからも選べます。')
    .build();
  sheet.getRange(startRow, col, rows, 1).setDataValidation(rule);
  sheet.getRange(startRow, col, rows, 1).setNumberFormat('yyyy"年"m"月"d"日"');
}

/** 客先クレーム管理台帳(CC)。2026-08-18リニューアル、フィールド構成は実機のxlsxダウンロードから調査したもの */
function buildCcLedgerSheetV2_(ss, masters, dataLists, sheetName) {
  var itemNames = DEFECT_ITEMS.map(function (item) { return item.name; });
  var recordFields = [
    { header: '台帳番号(CC)' },
    { header: '発生日(受領日)' },
    { header: '関連文書(客先からの連絡票等)' },
    { header: '関連文書(その他文書・ロット番号)' },
    { header: '客先名' },
    { header: '品番' },
    { header: '品名' },
    { header: 'クレーム内容分類', dropdown: itemNames },
    { header: 'クレーム内容詳細' },
    { header: 'NG数' },
    { header: '納入数' },
    { header: '加工者', dropdown: masters.kakosha },
    { header: '機械名', dropdown: masters.kishu },
    { header: '機械番号' },
    { header: '検査員', dropdown: dataLists.inspectors },
    { header: '出荷区分', dropdown: SHUKKA_KUBUN_CHOICES_ },
    { header: '単価' },
    { header: '金額' }
  ];
  return buildLedgerSheetV2_(ss, sheetName, recordFields, dataLists.staff,
    { qty: 'NG数', unitPrice: '単価', amount: '金額' });
}

/** 社内不良管理台帳(品証)(KP)。品証の検査で見つかった社内不良の改善計画書進捗を管理する(不良〇月とは別の台帳) */
function buildKpLedgerSheetV2_(ss, masters, dataLists, sheetName) {
  var itemNames = DEFECT_ITEMS.map(function (item) { return item.name; });
  var recordFields = [
    { header: '台帳番号(KP)' },
    { header: '発生日(受領日)' },
    { header: '関連文書(その他文書(客先等))' },
    { header: '客先名' },
    { header: '品番' },
    { header: '品名' },
    { header: 'クレーム内容分類', dropdown: itemNames },
    { header: 'クレーム内容詳細' },
    { header: '不良数' },
    { header: '加工数' },
    { header: '加工者', dropdown: masters.kakosha },
    { header: '機械名', dropdown: masters.kishu },
    { header: '機械番号' },
    { header: '検査員', dropdown: dataLists.inspectors },
    { header: '単価' },
    { header: '合計' }
  ];
  return buildLedgerSheetV2_(ss, sheetName, recordFields, dataLists.staff,
    { qty: '不良数', unitPrice: '単価', amount: '合計' });
}

/** 社内不良管理台帳(製造工程)(SK)。加工者自身がその場で申告する工程内不良の改善計画書進捗を管理する */
function buildSkLedgerSheetV2_(ss, masters, dataLists, sheetName) {
  var itemNames = DEFECT_ITEMS.map(function (item) { return item.name; });
  var recordFields = [
    { header: '台帳番号(SK)' },
    { header: '発生日(品質会議開催日)' },
    { header: '関連文書(製造工程不良一覧)' },
    { header: '関連文書(その他文書(客先等))' },
    { header: '客先名' },
    { header: '品番' },
    { header: '品名' },
    { header: '不良層別区分', dropdown: itemNames },
    { header: '不良詳細' },
    { header: '不良数' },
    { header: '加工者', dropdown: masters.kakosha },
    { header: '機械名', dropdown: masters.kishu },
    { header: '機械番号' }
  ];
  return buildLedgerSheetV2_(ss, sheetName, recordFields, dataLists.staff, null);
}

/**
 * 【1回だけ手動実行】CC/KP/SK台帳のリニューアル版を「（新）」を付けたシート名で追加する(2026-08-18)。
 * 現在稼働中の本物の台帳(客先クレーム管理台帳(CC)等)には一切触れない。内容を確認し、
 * 問題なければ swapToRedesignedLedgers で正式名称に切り替える。
 */
function addRedesignedLedgers() {
  var ss = getCurrentYearSpreadsheet_();
  var masters = fetchOrgMasterLists_();
  var dataLists = fetchDataSheetLists_(ss);
  buildCcLedgerSheetV2_(ss, masters, dataLists, CC_LEDGER_SHEET_NAME + '（新）');
  buildKpLedgerSheetV2_(ss, masters, dataLists, '社内不良管理台帳(品証)(KP)（新）');
  buildSkLedgerSheetV2_(ss, masters, dataLists, SK_LEDGER_SHEET_NAME + '（新）');
  SpreadsheetApp.flush();
  Logger.log('リニューアル版の3台帳を「（新）」付きシート名で追加しました。内容を確認し、問題なければ swapToRedesignedLedgers を実行してください。');
}

/**
 * 【1回だけ手動実行・要ユーザー確認後】addRedesignedLedgersで作った「（新）」付き台帳を正式名称に
 * 切り替える(2026-08-18)。現行の本物の台帳は削除せず「（旧）」を付けて非表示のまま残す
 * (中身を確認してから、不要ならユーザーが手動で削除する)。
 */
function swapToRedesignedLedgers() {
  var ss = getCurrentYearSpreadsheet_();
  var pairs = [
    { officialName: CC_LEDGER_SHEET_NAME },
    { officialName: '社内不良管理台帳(品証)(KP)' },
    { officialName: SK_LEDGER_SHEET_NAME }
  ];
  var log = [];
  pairs.forEach(function (p) {
    var newSheet = ss.getSheetByName(p.officialName + '（新）');
    if (!newSheet) { log.push(p.officialName + ': 「（新）」シートが見つからないためスキップ(先にaddRedesignedLedgersを実行してください)'); return; }
    var oldSheet = ss.getSheetByName(p.officialName);
    if (oldSheet) {
      oldSheet.setName(p.officialName + '（旧）');
      oldSheet.hideSheet();
    }
    newSheet.setName(p.officialName);
    log.push(p.officialName + ': 切り替えました' + (oldSheet ? '(旧シートは「' + p.officialName + '（旧）」として非表示で残しています)' : ''));
  });
  Logger.log(log.join('\n'));
}

/**
 * 【1回だけ手動実行】旧デザインのコード生成「仮」台帳3枚(客先クレーム管理台帳(CC)仮 等)を削除する。
 * リニューアル版(buildLedgerSheetV2_ベース)を導入したため、旧デザインの仮シートは不要と判断
 * (2026-08-18、ユーザー確認済み)。setupQualityDefectSystemFor_は今後もリニューアル版の仮シートを
 * 自動生成するため、削除しても支障はない。
 */
function removeOldTemporaryLedgerSheets() {
  var ss = getCurrentYearSpreadsheet_();
  var names = ['客先クレーム管理台帳(CC)仮', '社内不良管理台帳(品証)(KP)(仮)', '社内不良管理台帳(製造工程)(SK)(仮)'];
  var removed = [];
  names.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) { ss.deleteSheet(sheet); removed.push(name); }
  });
  Logger.log(removed.length > 0 ? '削除しました: ' + removed.join('、') : '該当するシートが見つかりませんでした。');
}

/** 既存の同名シートを削除してから新規作成する */
function replaceSheet_(ss, name) {
  var existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  return ss.insertSheet(name);
}

/** 列番号(1始まり)をA1形式の列名に変換 */
function columnToLetter_(column) {
  var letter = '';
  while (column > 0) {
    var remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}
