/**
 * 6月・7月分の過去データ移行(旧システム → 品質不具合管理システムの「不良〇月」への一括移行、1回限り)
 *
 * 【実行方法】
 * 1. このファイルを SetupSpreadsheet.gs・WebApi.gs と同じGASプロジェクト(品質不具合管理システムに紐づくプロジェクト)へ
 *    貼り付ける(DEFECT_ITEMSなど、SetupSpreadsheet.gs側の定義をそのまま使うため)
 * 2. まず「previewJuneJulyMigration」を実行し、実行ログ(表示 > ログ、またはCtrl+Enter)で
 *    「何件を、どう変換するか」を必ず確認する。ログの内容(特に警告行)をユーザーへ共有し、問題なければ次へ進む
 * 3. 問題なければ「migrateJuneJulyData」を実行する(実際に「不良6月」「不良7月」シートへ書き込む)
 *
 * 【移行元】
 * - 客先クレーム(CC)・社内不良(KP)管理台帳 2026年度 の「社内不具合6月」「社内不具合7月」シート(社内不良(KP)分)
 * - 品証さん差し戻し品の不良数入力 の「差し戻し品一覧表」(差し戻し分、年度=2026・月=6または7の行のみ)
 *
 * 【移行時に踏襲した割り切り(要確認)】
 * - 差し戻し側の旧シートには「加工数」に相当する列が無いため、移行後の加工数(I列)は空欄のままになる。
 *   正確な加工数が分かる場合は、移行後に手動で追記すること(空欄のままだと良品数の自動計算が
 *   0-不良数計のマイナス値のまま表示される)。
 * - 差し戻し側の不良項目名が新マスタ(DEFECT_ITEMS、29項目)に一致しない場合は、系統の近いカテゴリに
 *   丸め、元の項目名と数量を「不良項目詳細」列に残す(例: 「(旧項目名:圧痕修正不可) 5個」)。
 * - 品証担当者(C列)は、旧データの「検査担当者」(差し戻し)・「担当」(KP)をそのまま文字列で転記する
 *   (新システムのようなログイン認証済みの氏名ではなく、旧データの表記そのまま)。
 * - 単価は、旧データに対応する情報が無いため空欄のまま。担当者2は、KP側で2件目の不良項目に
 *   対応する担当者名が旧データにある場合のみ転記する(差し戻し側・1件目の不良項目は空欄)。
 * - キズ原因(U列)は、KP側の旧Googleフォームの選択値(例:「B.チャック内の切粉による圧痕」)が
 *   残っている場合、プレフィックス(「B.」等)を除いて新マスタ(KP_CAUSE_ITEMS)の名称と一致すれば転記する。
 * - 備考の先頭に「[旧データ移行]」と付記し、後から見分けられるようにする。書き込んだ全行(追加行含む)の
 *   送信ID列(V列)には`MIGRATE_MARKER`を入れており、`undoJuneJulyMigration`でやり直す際の目印にしている。
 */

var MIGRATE_TARGET_SS_ID = '1xeEybU6fhqvAEM02eEkUwBdOA_mD-fMUqKoCmckUcM0'; // 品質不具合管理システム 2026年度
var MIGRATE_KP_SOURCE_SS_ID = '1JRnbkwY5VF93dLTpyl2yRvaz5XOom6vM4_C6x54o5cM'; // 旧システムの2026年度ファイル(退役済み。2026-08-13にファイル名末尾へ「(旧アーカイブ)」を付記、IDは不変)
var MIGRATE_REWORK_SOURCE_SS_ID = '1xhBS77e1Jk3tJUSlNBmmIqHGlWXNl7fibQgqKaTPtFQ'; // 品証さん差し戻し品の不良数入力
var MIGRATE_MONTHS = [6, 7];
var MIGRATE_DATA_END_ROW = 115;
var MIGRATE_MARKER = 'MIGRATED_20260811'; // このスクリプトで書き込んだ行の送信ID列(V)に入れる目印

/** 何も書き込まず、移行対象を集計してログに出すだけの確認用(必ず先にこちらを実行する) */
function previewJuneJulyMigration() {
  runJuneJulyMigration_(true);
}

/**
 * 以前 migrateJuneJulyData で書き込んだ行(送信ID列がMIGRATE_MARKERの行)をすべてクリアする。
 * キズ原因の転記漏れなど、移行内容をやり直したい場合にこれを実行してから migrateJuneJulyData を再実行する。
 * (MIGRATE_MARKERを書き込む前の版で移行した行は見つけられないので、その場合はclearMonthSheetDataForRedoを使う)
 */
function undoJuneJulyMigration() {
  var targetSs = SpreadsheetApp.openById(MIGRATE_TARGET_SS_ID);
  var log = [];
  MIGRATE_MONTHS.forEach(function (month) {
    var sheet = targetSs.getSheetByName('不良' + month + '月');
    if (!sheet) return;
    var values = sheet.getRange(2, 1, MIGRATE_DATA_END_ROW - 1, 22).getValues();
    var cleared = 0;
    for (var i = 0; i < values.length; i++) {
      if (values[i][21] !== MIGRATE_MARKER) continue; // V列(22列目、0始まりで21)
      var row = i + 2;
      sheet.getRange(row, 1, 1, 9).clearContent();   // A-I
      sheet.getRange(row, 11, 1, 5).clearContent();  // K-O
      sheet.getRange(row, 16).clearContent();        // P 単価
      sheet.getRange(row, 18, 1, 2).clearContent();  // R-S 備考・材質
      sheet.getRange(row, 21, 1, 2).clearContent();  // U-V キズ原因・送信ID
      cleared++;
    }
    log.push('不良' + month + '月: ' + cleared + '行をクリアしました');
  });
  Logger.log(log.join('\n'));
}

/**
 * 「不良6月」「不良7月」のデータ行(2〜115行)を丸ごとクリアする(J・Q・T列の自動計算式は残す)。
 * MIGRATE_MARKERを書き込む前の版でmigrateJuneJulyDataを実行してしまった場合など、
 * undoJuneJulyMigrationで対象行を見つけられない時にこちらを使う。
 * 【注意】この2シートに今回の移行以外のデータが無いことを確認してから実行すること。
 */
function clearMonthSheetDataForRedo() {
  var targetSs = SpreadsheetApp.openById(MIGRATE_TARGET_SS_ID);
  var log = [];
  MIGRATE_MONTHS.forEach(function (month) {
    var sheet = targetSs.getSheetByName('不良' + month + '月');
    if (!sheet) return;
    var rows = MIGRATE_DATA_END_ROW - 1;
    sheet.getRange(2, 1, rows, 9).clearContent();   // A-I
    sheet.getRange(2, 11, rows, 5).clearContent();  // K-O
    sheet.getRange(2, 16, rows, 1).clearContent();  // P 単価
    sheet.getRange(2, 18, rows, 2).clearContent();  // R-S 備考・材質
    sheet.getRange(2, 21, rows, 2).clearContent();  // U-V キズ原因・送信ID
    log.push('不良' + month + '月: データ行をクリアしました');
  });
  Logger.log(log.join('\n'));
}

/**
 * 【調査用・書き込みなし】「不良6月」「不良7月」の各行を、送信ID列(V列)が MIGRATE_MARKER かどうかで
 * 仕分けてログに出す。clearMonthSheetDataForRedoを実行せずにmigrateJuneJulyDataを2回実行してしまった
 * 場合、1回目(印なし)の行がそのまま残り、2回目(印あり)の行が重複して増える現象が起きる。
 * これを実行して、印なし行(=消してよい古いデータ)が本当に重複と一致するか確認してから
 * removeDuplicateMigratedRowsを実行すること。
 */
function previewDuplicateCleanup() {
  var targetSs = SpreadsheetApp.openById(MIGRATE_TARGET_SS_ID);
  var log = [];
  var rows = MIGRATE_DATA_END_ROW - 1;

  MIGRATE_MONTHS.forEach(function (month) {
    var sheet = targetSs.getSheetByName('不良' + month + '月');
    if (!sheet) return;
    var values = sheet.getRange(2, 1, rows, 22).getValues();

    var markedCount = 0, unmarkedCount = 0;
    var unmarkedSample = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (!row[0] && !row[11]) continue; // A列・L列とも空欄は無関係の行
      if (row[21] === MIGRATE_MARKER) {
        markedCount++;
      } else {
        unmarkedCount++;
        if (unmarkedSample.length < 5) {
          unmarkedSample.push('行' + (i + 2) + ': ' + row[3] + ' / ' + row[4] + ' / ' + row[11]);
        }
      }
    }
    log.push('不良' + month + '月: 印あり(MIGRATE_MARKER) ' + markedCount + '行 / 印なし ' + unmarkedCount + '行');
    if (unmarkedSample.length > 0) {
      log.push('  印なし行の例(得意先/品番/不良項目):\n    ' + unmarkedSample.join('\n    '));
    }
  });

  Logger.log(log.join('\n'));
}

/**
 * previewDuplicateCleanupで「印なし」行が重複(古い1回目の移行データ)だと確認できた場合に実行する。
 * 印なし(A列・L列のどちらかが埋まっている)行の内容だけをクリアし、印あり(MIGRATE_MARKER)行だけを
 * 上から詰め直す(行削除ではなくセルのクリア+書き直しなので、J・Q・T列の自動計算式やシートの行数構成は
 * 壊さない)。最後にaddRecordBordersToAllMonths(SetupSpreadsheet.gs)を再実行して枠線を引き直すこと。
 */
function removeDuplicateMigratedRows() {
  var targetSs = SpreadsheetApp.openById(MIGRATE_TARGET_SS_ID);
  var log = [];
  var rows = MIGRATE_DATA_END_ROW - 1;

  MIGRATE_MONTHS.forEach(function (month) {
    var sheet = targetSs.getSheetByName('不良' + month + '月');
    if (!sheet) return;
    var values = sheet.getRange(2, 1, rows, 22).getValues();

    // MIGRATE_MARKERが付いている行(=修正後の正しいデータ)だけを残す
    var keepRows = values.filter(function (row) {
      return row[21] === MIGRATE_MARKER;
    });
    var removedCount = values.filter(function (row) {
      return (row[0] || row[11]) && row[21] !== MIGRATE_MARKER;
    }).length;

    // 一旦データ行を全部クリアしてから、残すべき行だけを上から詰めて書き直す
    sheet.getRange(2, 1, rows, 9).clearContent();   // A-I
    sheet.getRange(2, 11, rows, 5).clearContent();  // K-O
    sheet.getRange(2, 16, rows, 1).clearContent();  // P 単価
    sheet.getRange(2, 18, rows, 2).clearContent();  // R-S 備考・材質
    sheet.getRange(2, 21, rows, 2).clearContent();  // U-V キズ原因・送信ID

    // J・Q・T列(10・17・20列目)は数式セルなので、getValuesで拾った計算結果を書き戻さないよう
    // 列グループごとに分けて書く(migrateJuneJulyDataのwriteRecordsToSheet_と同じ列の切り方)
    if (keepRows.length > 0) {
      var n = keepRows.length;
      sheet.getRange(2, 1, n, 9).setValues(keepRows.map(function (r) { return r.slice(0, 9); }));    // A-I
      sheet.getRange(2, 11, n, 5).setValues(keepRows.map(function (r) { return r.slice(10, 15); }));  // K-O
      sheet.getRange(2, 16, n, 1).setValues(keepRows.map(function (r) { return [r[15]]; }));          // P 単価
      sheet.getRange(2, 18, n, 2).setValues(keepRows.map(function (r) { return r.slice(17, 19); }));  // R-S
      sheet.getRange(2, 21, n, 2).setValues(keepRows.map(function (r) { return r.slice(20, 22); }));  // U-V
    }

    log.push('不良' + month + '月: 重複' + removedCount + '行を削除、' + keepRows.length + '行を残しました');
  });

  Logger.log(log.join('\n') + '\n\n完了後、SetupSpreadsheet.gsの addRecordBordersToAllMonths を再実行して枠線を引き直してください。');
}

/**
 * 【調査用・書き込みなし】旧「社内不具合〇月」シートのG列(不具合内容1)に設定されている
 * プルダウン(データの入力規則)の選択肢をそのままログに出す。月によって選択肢が違わないか、
 * 12ヶ月すべてを確認する。不良〇月のL列プルダウンをこれに合わせるための事前調査用。
 */
function debugInspectGColumnValidation() {
  var ss = SpreadsheetApp.openById(MIGRATE_KP_SOURCE_SS_ID);
  var log = [];
  var ALL_MONTHS = [6, 7, 8, 9, 10, 11, 12, 1, 2, 3, 4, 5];
  var firstList = null;

  ALL_MONTHS.forEach(function (month) {
    var sheetName = '社内不具合' + month + '月';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      log.push('【' + sheetName + '】シートが見つかりません');
      return;
    }
    var values = sheet.getRange(1, 1, Math.min(10, sheet.getLastRow()), sheet.getLastColumn()).getValues();
    var headerRowIndex = -1;
    for (var hr = 0; hr < values.length; hr++) {
      if (values[hr].indexOf('図番') !== -1) { headerRowIndex = hr; break; }
    }
    if (headerRowIndex === -1) {
      log.push('【' + sheetName + '】ヘッダー行が見つかりません');
      return;
    }
    var colItem1 = values[headerRowIndex].indexOf('不具合内容1') + 1; // 1始まりの列番号に変換
    if (colItem1 === 0) {
      log.push('【' + sheetName + '】不具合内容1列が見つかりません');
      return;
    }
    var dataRow = headerRowIndex + 2; // ヘッダーの次の行(1始まり)
    var validation = sheet.getRange(dataRow, colItem1).getDataValidation();
    if (!validation) {
      log.push('【' + sheetName + '】G列(' + dataRow + '行目)にプルダウンの設定がありません');
      return;
    }
    var criteriaType = validation.getCriteriaType();
    var values2 = validation.getCriteriaValues();
    log.push('【' + sheetName + '】criteriaType=' + criteriaType);
    if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var list = values2[0];
      log.push('  選択肢(' + list.length + '件): ' + JSON.stringify(list));
      if (!firstList) firstList = { month: month, list: list };
      else if (JSON.stringify(list) !== JSON.stringify(firstList.list)) {
        log.push('  ※' + firstList.month + '月と選択肢が異なります');
      }
    } else if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      var range = values2[0];
      log.push('  範囲参照: ' + range.getSheet().getName() + '!' + range.getA1Notation());
      var rangeValues = range.getValues().map(function (r) { return r[0]; }).filter(function (v) { return v; });
      log.push('  範囲の中身(' + rangeValues.length + '件): ' + JSON.stringify(rangeValues));
    } else {
      log.push('  (リスト形式ではありません: ' + JSON.stringify(values2) + ')');
    }
  });

  Logger.log(log.join('\n'));
}

/**
 * 【調査用・書き込みなし】旧スプレッドシート(客先クレーム(CC)・社内不良(KP)管理台帳 2026年度)の
 * 全シートに埋め込まれている棒グラフ・折れ線グラフ等を洗い出し、種類・タイトル・参照範囲を
 * そのままログに出す。新システムにも同等(できれば改善した)グラフを作るための事前調査用。
 */
function debugInspectCharts() {
  var ss = SpreadsheetApp.openById(MIGRATE_KP_SOURCE_SS_ID);
  var log = [];

  ss.getSheets().forEach(function (sheet) {
    var charts = sheet.getCharts();
    if (charts.length === 0) return;
    log.push('【シート「' + sheet.getName() + '」】グラフ' + charts.length + '件');
    charts.forEach(function (chart, i) {
      var info = chart.getContainerInfo();
      var options = chart.getOptions();
      log.push('  ' + (i + 1) + '. 種類=' + options.get('chartType') +
        ' / タイトル=' + (options.get('title') || '(なし)') +
        ' / 位置=行' + info.getAnchorRow() + ' 列' + info.getAnchorColumn());
      var ranges = chart.getRanges();
      ranges.forEach(function (r) {
        log.push('     参照範囲: ' + r.getSheet().getName() + '!' + r.getA1Notation());
      });
    });
  });

  if (log.length === 0) {
    log.push('グラフが見つかりませんでした(全' + ss.getSheets().length + 'シートを確認)。');
  }
  Logger.log(log.join('\n'));
}

/**
 * 【調査用・書き込みなし】旧スプレッドシートの全グラフを画像(PNG)として書き出し、
 * マイドライブに新規フォルダを作って保存する。APIからは種類(棒/折れ線等)が正しく
 * 取得できなかったため、実際の見た目を画像で確認するための代替手段。
 * 実行後、ログに出るフォルダのURLをそのまま共有してください。
 */
function exportChartsAsImages() {
  var ss = SpreadsheetApp.openById(MIGRATE_KP_SOURCE_SS_ID);
  var folder = DriveApp.createFolder('グラフ確認用_' + new Date().getTime());
  var log = [];
  var count = 0;

  ss.getSheets().forEach(function (sheet) {
    var charts = sheet.getCharts();
    charts.forEach(function (chart) {
      count++;
      var title = (chart.getOptions().get('title') || ('グラフ' + count)).toString();
      var fileName = count + '_' + title.replace(/[\\\/:*?"<>|]/g, '_') + '.png';
      try {
        var blob = chart.getAs('image/png').setName(fileName);
        folder.createFile(blob);
        log.push('OK: ' + fileName);
      } catch (err) {
        log.push('失敗: ' + fileName + ' (' + err.message + ')');
      }
    });
  });

  Logger.log('フォルダ「' + folder.getName() + '」に' + count + '件の画像を保存しました。\nURL: ' + folder.getUrl() + '\n\n' + log.join('\n'));
}

/**
 * 【調査用・書き込みなし】本物のCC/KP/SK台帳3枚(客先クレーム管理台帳(CC)・社内不良管理台帳(品証)(KP)・
 * 社内不良管理台帳(製造工程)(SK))の見出し行を走査し、「日」を含む見出し(発生日・完了日等)の列について、
 * データ行1行目の入力規則(プルダウン設定の有無)・表示形式・結合状態をログに出す(2026-08-17新設)。
 * ユーザーから「年月日の入力がプルダウンになっているのを、ダブルクリックでカレンダー入力できる
 * ようにしたい」との依頼を受け、現状の設定(本当にリスト形式のプルダウンなのか、単なる日付書式なのか)
 * を実物で確認してから対応方針を決めるための調査用。
 */
function debugInspectDateColumnValidation() {
  var ss = getCurrentYearSpreadsheet_();
  var sheetNames = [CC_LEDGER_SHEET_NAME, '社内不良管理台帳(品証)(KP)', '社内不良管理台帳(製造工程)(SK)'];
  var log = [];

  sheetNames.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { log.push('【' + name + '】シートが見つかりません'); return; }

    var frozen = sheet.getFrozenRows();
    var lastCol = sheet.getLastColumn();
    var dataRow = frozen + 1;
    log.push('【' + name + '】固定行数=' + frozen + '(データ開始行=' + dataRow + ')、最終列=' + lastCol);

    var headerRows = sheet.getRange(1, 1, Math.max(frozen, 1), lastCol).getValues();
    var found = 0;
    for (var r = 0; r < headerRows.length; r++) {
      for (var c = 0; c < headerRows[r].length; c++) {
        var text = headerRows[r][c] ? headerRows[r][c].toString().trim() : '';
        if (!text || text.indexOf('日') === -1 || text.length > 25) continue;
        found++;
        var col = c + 1;
        var cell = sheet.getRange(dataRow, col);
        var validation = cell.getDataValidation();
        var vDesc = '設定なし';
        if (validation) {
          var ct = validation.getCriteriaType();
          var cv = validation.getCriteriaValues();
          if (ct === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
            vDesc = 'リスト直接指定(' + cv[0].length + '件、プルダウン): ' + JSON.stringify(cv[0].slice(0, 5));
          } else if (ct === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
            vDesc = '範囲参照リスト(プルダウン): ' + cv[0].getSheet().getName() + '!' + cv[0].getA1Notation();
          } else if (ct === SpreadsheetApp.DataValidationCriteria.DATE_IS_VALID_DATE) {
            vDesc = '日付検証(カレンダー入力可)';
          } else {
            vDesc = ct + ': ' + JSON.stringify(cv);
          }
        }
        var numFmt = cell.getNumberFormat();
        var isMerged = sheet.getRange(r + 1, col).isPartOfMerge();
        log.push('  ' + columnToLetter_(col) + '列(見出し行' + (r + 1) + '「' + text + '」): 入力規則=' + vDesc +
          ' / 表示形式=' + numFmt + ' / 見出しセルは結合=' + isMerged);
      }
    }
    if (found === 0) log.push('  「日」を含む見出しが見つかりませんでした。');
  });

  Logger.log(log.join('\n'));
}

/**
 * 【調査用・書き込みなし】旧スプレッドシートの「グラフ」シートのうち、指定した2範囲
 * (B78:U91、B65:AD71)をそのままログに出す。「クレーム集計」シート新設の依頼を受けて、
 * 旧システムのこの範囲に何が入っているか(表の構造・見出し・結合セル)を実物で確認するための調査用。
 */
function debugInspectClaimSummaryRanges() {
  var ss = SpreadsheetApp.openById(MIGRATE_KP_SOURCE_SS_ID);
  var sheet = ss.getSheetByName('グラフ');
  if (!sheet) {
    Logger.log('「グラフ」シートが見つかりません(全シート: ' + ss.getSheets().map(function (s) { return s.getName(); }).join(', ') + ')');
    return;
  }

  var log = [];
  ['B78:U91', 'B65:AD71'].forEach(function (a1) {
    log.push('===== ' + a1 + ' =====');
    var range = sheet.getRange(a1);
    var startRow = range.getRow();
    var values = range.getValues();
    var merges = range.getMergedRanges().map(function (m) { return m.getA1Notation(); });
    log.push('結合セル: ' + (merges.length ? merges.join(', ') : 'なし'));
    values.forEach(function (row, i) {
      var line = row.map(function (v) { return (v === '' || v === null) ? '・' : v; }).join(' | ');
      log.push((startRow + i) + '行目: ' + line);
    });
    log.push('');
  });

  Logger.log(log.join('\n'));
}

/**
 * 【調査用・書き込みなし】新システム側(現行年度)の「客先クレーム管理台帳(CC)」(ユーザーが
 * 旧システム形式で手作りした本物のシート)の実際のヘッダー行・列位置・データ数行、および
 * 「データ」シートの「クレーム内容分類」関連プルダウン候補をログに出す。
 * 「クレーム集計」シート新設にあたり、集計元となるこのシートの列構成(特にクレーム内容分類の
 * 実際の選択肢)が新マスタ(DEFECT_ITEMS、40項目)と同じか、旧システムのより粗い分類のままかを
 * 確認するための調査用。
 */
function debugInspectCcLedgerColumns() {
  var log = [];
  var ss = getCurrentYearSpreadsheet_();
  var sheet = ss.getSheetByName('客先クレーム管理台帳(CC)');
  if (!sheet) {
    log.push('「客先クレーム管理台帳(CC)」シートが見つかりません(全シート: ' + ss.getSheets().map(function (s) { return s.getName(); }).join(', ') + ')');
    Logger.log(log.join('\n'));
    return;
  }

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  log.push('シート「' + sheet.getName() + '」: ' + lastRow + '行 × ' + lastCol + '列(見た目上の最終行列)');

  log.push('--- 先頭5行(A〜' + columnToLetter_(lastCol) + '列) ---');
  var headRows = sheet.getRange(1, 1, Math.min(5, lastRow), lastCol).getValues();
  headRows.forEach(function (row, i) {
    log.push((i + 1) + '行目: ' + row.map(function (v) { return (v === '' || v === null) ? '・' : v; }).join(' | '));
  });

  // ヘッダーらしき行(「クレーム内容分類」を含む行)を探し、その列のデータ入力規則とデータ側の実データを見る
  var headerRowIndex = -1, colIndex = -1;
  for (var r = 0; r < headRows.length; r++) {
    var idx = headRows[r].indexOf('クレーム内容分類');
    if (idx !== -1) { headerRowIndex = r; colIndex = idx + 1; break; }
  }
  if (colIndex !== -1) {
    log.push('--- 「クレーム内容分類」列: ' + columnToLetter_(colIndex) + '列(' + colIndex + '列目)、ヘッダー行=' + (headerRowIndex + 1) + '行目 ---');
    var dataStartRow = headerRowIndex + 2;
    if (dataStartRow <= lastRow) {
      var validation = sheet.getRange(dataStartRow, colIndex).getDataValidation();
      if (validation) {
        var ct = validation.getCriteriaType();
        var cv = validation.getCriteriaValues();
        if (ct === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          log.push('入力規則(直接リスト、' + cv[0].length + '件): ' + JSON.stringify(cv[0]));
        } else if (ct === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          var range = cv[0];
          var listVals = range.getValues().map(function (r) { return r[0]; }).filter(function (v) { return v; });
          log.push('入力規則(範囲参照 ' + range.getSheet().getName() + '!' + range.getA1Notation() + '、' + listVals.length + '件): ' + JSON.stringify(listVals));
        } else {
          log.push('入力規則: リスト形式ではありません(' + JSON.stringify(cv) + ')');
        }
      } else {
        log.push('入力規則: 設定なし');
      }
      var actualEnd = Math.min(dataStartRow + 14, lastRow);
      if (actualEnd >= dataStartRow) {
        var actualVals = sheet.getRange(dataStartRow, colIndex, actualEnd - dataStartRow + 1, 1).getValues().map(function (r) { return r[0]; }).filter(function (v) { return v; });
        log.push('実際に入力済みの値(先頭' + actualVals.length + '件、重複含む): ' + JSON.stringify(actualVals));
      }
    }
  } else {
    log.push('「クレーム内容分類」という見出しが先頭5行に見つかりませんでした。');
  }

  // データシート側の「クレーム内容分類」関連の列も確認(旧システムでは「クレーム内容分類(CC・KP用)」という見出しがあった)
  var dataSheet = ss.getSheetByName('データ');
  if (dataSheet) {
    var dLastCol = dataSheet.getLastColumn();
    var dHeaderRow = dataSheet.getRange(1, 1, 1, dLastCol).getValues()[0];
    log.push('--- 「データ」シートのヘッダー行(1行目) ---');
    log.push(dHeaderRow.map(function (v, i) { return columnToLetter_(i + 1) + ':' + (v || '・'); }).join(' | '));
    for (var c = 0; c < dHeaderRow.length; c++) {
      if (dHeaderRow[c] && dHeaderRow[c].toString().indexOf('クレーム') !== -1) {
        var colLetter = columnToLetter_(c + 1);
        var lastR = dataSheet.getLastRow();
        var vals = dataSheet.getRange(2, c + 1, Math.max(lastR - 1, 1), 1).getValues().map(function (r) { return r[0]; }).filter(function (v) { return v; });
        log.push(colLetter + '列「' + dHeaderRow[c] + '」の中身(' + vals.length + '件): ' + JSON.stringify(vals));
      }
    }
  } else {
    log.push('「データ」シートが見つかりません。');
  }

  Logger.log(log.join('\n'));
}

/**
 * 【調査用】KP側「社内不具合6月」「社内不具合7月」シートの実際のヘッダー行(先頭10行)と
 * 年度2026の差し戻しデータの有無を、そのままログに出す(列名の推測が外れたため実物を確認する用)
 */
function debugInspectSources() {
  var log = [];
  var kpSs = SpreadsheetApp.openById(MIGRATE_KP_SOURCE_SS_ID);

  MIGRATE_MONTHS.forEach(function (month) {
    var sheetName = '社内不具合' + month + '月';
    var sheet = kpSs.getSheetByName(sheetName);
    if (!sheet) {
      log.push('【' + sheetName + '】シートが見つかりません');
      return;
    }
    log.push('【' + sheetName + '】先頭10行:');
    var values = sheet.getRange(1, 1, Math.min(10, sheet.getLastRow()), sheet.getLastColumn()).getValues();
    values.forEach(function (row, i) {
      log.push('  行' + (i + 1) + ': ' + JSON.stringify(row));
    });
  });

  var reworkSs = SpreadsheetApp.openById(MIGRATE_REWORK_SOURCE_SS_ID);
  reworkSs.getSheets().forEach(function (sheet) {
    var values = sheet.getDataRange().getValues();
    for (var r = 0; r < values.length; r++) {
      if (values[r].indexOf('差し戻し日') === -1) continue;
      var header = values[r];
      var colYear = header.indexOf('年度');
      var colMonth = header.indexOf('月');
      log.push('【差し戻し(' + sheet.getName() + ' 行' + (r + 1) + ')】年度列=' + colYear + ' 月列=' + colMonth);
      // このヘッダーブロックの年度・月の値をユニークに集計(実際に入っている値を確認するため)
      var yearMonthSet = {};
      for (var d = r + 1; d < values.length; d++) {
        var row = values[d];
        if (colYear === -1 || !row[colYear]) continue;
        var key = row[colYear] + '/' + row[colMonth];
        yearMonthSet[key] = (yearMonthSet[key] || 0) + 1;
      }
      log.push('    年度/月ごとの件数: ' + JSON.stringify(yearMonthSet));
    }
  });

  Logger.log(log.join('\n'));
}

/** 実際に「不良6月」「不良7月」へ書き込む(previewJuneJulyMigrationで内容を確認した後に実行する) */
function migrateJuneJulyData() {
  runJuneJulyMigration_(false);
}

function runJuneJulyMigration_(dryRun) {
  var targetSs = SpreadsheetApp.openById(MIGRATE_TARGET_SS_ID);
  var log = [];

  MIGRATE_MONTHS.forEach(function (month) {
    var targetSheet = targetSs.getSheetByName('不良' + month + '月');
    if (!targetSheet) {
      log.push('【エラー】不良' + month + '月 シートが見つかりません。スキップします。');
      return;
    }

    var kpRecords = extractKpRecords_(month, log);
    var reworkRecords = extractReworkRecords_(month, log);
    var records = kpRecords.concat(reworkRecords);

    log.push('--- ' + month + '月: KP ' + kpRecords.length + '件、差し戻し ' + reworkRecords.length + '件 ---');
    records.forEach(function (rec) {
      log.push('  [' + rec.shochiKubun + '] ' + rec.customer + ' / ' + rec.drawing + ' / 担当:' + rec.worker +
        ' / 品証:' + rec.quality + ' / キズ原因:' + (rec.kizuGenin || '(なし)') +
        ' / 項目:' + rec.items.map(function (it) { return it.name + '×' + it.qty; }).join(','));
    });

    if (!dryRun) {
      var written = writeRecordsToSheet_(targetSheet, records);
      log.push('  → ' + written + '行を書き込みました');
    }
  });

  Logger.log((dryRun ? '【プレビューのみ・書き込みなし】\n' : '【書き込み完了】\n') + log.join('\n'));
}

/** 「社内不具合6月」「社内不具合7月」から社内不良(KP)分のレコードを抽出する */
function extractKpRecords_(month, log) {
  var ss = SpreadsheetApp.openById(MIGRATE_KP_SOURCE_SS_ID);
  var sheetName = '社内不具合' + month + '月';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    log.push('【警告】KP側「' + sheetName + '」シートが見つかりません。KP分はスキップします。');
    return [];
  }

  var values = sheet.getDataRange().getValues();

  // ヘッダー行は先頭行とは限らない(結合セルのタイトル行が上にあることがある)ため、
  // 「図番」を含む行を先頭10行から探してヘッダー行とする
  var headerRowIndex = -1;
  for (var hr = 0; hr < Math.min(values.length, 10); hr++) {
    if (values[hr].indexOf('図番') !== -1) { headerRowIndex = hr; break; }
  }
  if (headerRowIndex === -1) {
    log.push('【警告】KP側「' + sheetName + '」で「図番」を含むヘッダー行が先頭10行に見つかりません。KP分はスキップします。');
    return [];
  }
  var header = values[headerRowIndex];
  var colDrawing = header.indexOf('図番');
  var colCustomer = header.indexOf('客先');
  var colQty = header.indexOf('加工数');
  var colItem1 = header.indexOf('不具合内容1');
  var colItem2 = header.indexOf('その他不具合内容');
  var colPrice = header.indexOf('単価');
  var colRemark = header.indexOf('備考');
  var colMaterial = header.indexOf('材質');
  var colDate = header.indexOf('日付');
  var colRate = header.indexOf('不良率');
  // 「不良率」の2つ後ろの列に、旧Googleフォームの「キズ原因」選択値(例:「B.チャック内の切粉による圧痕」)が
  // 見出し無しで入っている(旧コード.jsのS列=scratchRaw相当)。分類の精度を上げるための補助情報として使う。
  var colScratch = colRate !== -1 ? colRate + 2 : -1;

  if (colItem1 === -1) {
    log.push('【警告】KP側「' + sheetName + '」の列見出しが想定と違います(不具合内容1が見つからない)。KP分はスキップします。');
    return [];
  }

  // 1件の不良で不良項目が3つ以上ある場合、旧シートは図番が空欄の追加行(その他不具合内容の列だけ埋まっている)で
  // 表現しているため、図番が現れるまでの後続の空欄行を同じレコードの追加項目として拾う
  var records = [];
  var current = null;
  for (var r = headerRowIndex + 1; r < values.length; r++) {
    var row = values[r];

    if (row[colDrawing]) {
      // 新しいレコードの開始
      if (current) records.push(current);
      var scratchTag = colScratch !== -1 ? row[colScratch] : '';

      // 【重要】旧シートは「不具合内容1」(カテゴリ的な短い言い方)と「その他不具合内容」(詳細な言い方)が
      // 実は同じ1件の不良を2通りに表現しているだけで、数量(不具合数)もH列・K列に同じ値が二重に
      // 入る仕様だった(旧コード.jsのngQty1がH・Kの両方に書かれる作り)。そのためメイン行からは
      // 不良項目を1件だけ作る。分類の精度を上げるため、より具体的な「その他不具合内容」を優先して
      // 判定に使い、両方のテキストを不良項目詳細に残す。
      var items = [];
      var rawName1 = colItem1 !== -1 && row[colItem1] ? row[colItem1].toString() : '';
      var rawName2 = colItem2 !== -1 && row[colItem2] ? row[colItem2].toString() : '';
      if (rawName1 || rawName2) {
        var primaryName = rawName2 || rawName1;
        var detailText = rawName1 && rawName2 && rawName1 !== rawName2
          ? '(旧内容:' + rawName1 + '／' + rawName2 + ')'
          : '(旧内容:' + primaryName + ')';
        items.push({
          name: mapDefectItemName_(primaryName, scratchTag, log),
          qty: (colItem1 !== -1 ? row[colItem1 + 1] : 0) || 0,
          detail: detailText,
          worker2: ''
        });
      }
      current = {
        shochiKubun: '社内不良(KP)',
        timestamp: colDate !== -1 ? parseSourceDate_(row[colDate]) : new Date(),
        quality: '',
        customer: colCustomer !== -1 ? row[colCustomer] : '',
        drawing: row[colDrawing],
        worker: colItem1 !== -1 ? row[colItem1 + 2] : '', // 担当(不具合内容1の2つ後ろの列)
        kishu: '',
        setsubi: '',
        suryo: colQty !== -1 ? row[colQty] : '',
        price: colPrice !== -1 ? row[colPrice] : '',
        remark: '[旧データ移行]' + (colRemark !== -1 && row[colRemark] ? ' ' + row[colRemark] : ''),
        material: colMaterial !== -1 ? row[colMaterial] : '',
        kizuGenin: mapScratchToKizuGenin_(scratchTag),
        items: items
      };
    } else if (current && colItem2 !== -1 && row[colItem2]) {
      // 図番が空欄で、その他不具合内容だけ埋まっている行 = 直前レコードの3件目以降の不良項目
      var rawName3 = row[colItem2].toString();
      var scratchTag3 = colScratch !== -1 ? row[colScratch] : '';
      current.items.push({ name: mapDefectItemName_(rawName3, scratchTag3, log), qty: row[colItem2 + 1] || 0, detail: '(旧内容:' + rawName3 + ')', worker2: row[colItem2 + 2] || '' });
    }
  }
  if (current) records.push(current);

  // 不良項目が1つも無いレコードは除外
  return records.filter(function (rec) { return rec.items.length > 0; });
}

/** 「差し戻し品一覧表」から差し戻し分(年度2026・指定月)のレコードを抽出する */
function extractReworkRecords_(month, log) {
  var ss = SpreadsheetApp.openById(MIGRATE_REWORK_SOURCE_SS_ID);
  var sheets = ss.getSheets();
  var records = [];
  var seen = {}; // 同じ内容のブロックが複数箇所にある場合の重複防止
  var blockCount = 0;
  var rawRowCount = 0; // 年度・月一致で見つかった生の行数(丸め・重複排除前)

  sheets.forEach(function (sheet) {
    var values = sheet.getDataRange().getValues();
    for (var r = 0; r < values.length; r++) {
      if (values[r].indexOf('差し戻し日') === -1) continue;
      blockCount++;
      var block = extractReworkBlock_(values, r, month, log);
      rawRowCount += block.length;
      block.forEach(function (rec) {
        var key = JSON.stringify([rec.timestamp, rec.customer, rec.drawing, rec.worker, rec.items]);
        if (seen[key]) return;
        seen[key] = true;
        records.push(rec);
      });
    }
  });

  if (blockCount === 0) {
    log.push('【警告】差し戻し側で「差し戻し日」を含むヘッダー行が1つも見つかりませんでした(全シート走査)。差し戻し分はスキップします。');
  } else {
    log.push('  (差し戻し側: 「差し戻し日」ヘッダーを' + blockCount + '箇所検出、うち年度2026・' + month + '月の行が' + rawRowCount + '件、重複排除後' + records.length + '件)');
  }
  return records;
}

/** ヘッダー行(headerRowIndex)を起点に、指定月・年度2026のデータ行を読み取る */
function extractReworkBlock_(values, headerRowIndex, month, log) {
  var header = values[headerRowIndex];
  var colDate = header.indexOf('差し戻し日');
  var colWorker = header.indexOf('加工者氏名');
  var colCustomer = header.indexOf('得意先');
  var colDrawing = header.indexOf('図番');
  var colMaterial = header.indexOf('材料名');
  var colKishu = header.indexOf('機種名');
  var colSetsubi = header.indexOf('設備№');
  var colQuality = header.indexOf('検査担当者');
  var colYear = header.indexOf('年度');
  var colMonth = header.indexOf('月');
  var colDay = header.indexOf('日');
  var colRemark = header.indexOf('備考');

  if (colDate === -1 || colYear === -1 || colMonth === -1 || colQuality === -1) {
    log.push('【警告】差し戻し側のヘッダー行(行' + (headerRowIndex + 1) + ')の列見出しが想定と違うため、このブロックはスキップします。');
    return [];
  }

  // 不良項目の列は「検査担当者」の次〜「年度」の手前までの範囲とみなす(バージョンによって項目数が違うため)
  var itemColStart = colQuality + 1;
  var itemColEnd = colYear - 1;

  var records = [];
  for (var r = headerRowIndex + 1; r < values.length; r++) {
    var row = values[r];
    if (!row[colDrawing] && !row[colCustomer]) continue;
    var year = Number(row[colYear]);
    var mon = Number(row[colMonth]);
    if (year !== 2026 || mon !== month) continue;

    var items = [];
    for (var c = itemColStart; c <= itemColEnd; c++) {
      var qty = row[c];
      if (qty === '' || qty === null || !Number(qty)) continue;
      var name = header[c] ? header[c].toString() : '';
      items.push({ name: mapDefectItemName_(name, '', log), qty: Number(qty), detail: '(旧項目名:' + name + ')', worker2: '' });
    }
    if (items.length === 0) continue;

    var day = Number(row[colDay]) || 1;
    records.push({
      shochiKubun: '差し戻し',
      timestamp: new Date(year, mon - 1, day),
      quality: colQuality !== -1 ? row[colQuality] : '',
      customer: colCustomer !== -1 ? row[colCustomer] : '',
      drawing: colDrawing !== -1 ? row[colDrawing] : '',
      worker: colWorker !== -1 ? row[colWorker] : '',
      kishu: colKishu !== -1 ? row[colKishu] : '',
      setsubi: colSetsubi !== -1 ? row[colSetsubi] : '',
      suryo: '', // 旧シートに加工数の記録が無いため空欄(移行時の既知の割り切り、ファイル冒頭コメント参照)
      price: '',
      remark: '[旧データ移行・加工数不明]' + (colRemark !== -1 && row[colRemark] ? ' ' + row[colRemark] : ''),
      material: colMaterial !== -1 ? row[colMaterial] : '',
      kizuGenin: '', // 差し戻し側にはキズ原因に相当する記録が無い
      items: items
    });
  }
  return records;
}

// キーワード → 新マスタ項目名(2026-08-12、40項目マスタに合わせて全面書き換え)。優先順位順(先に一致した方を採用)。
// 新マスタには「圧痕」「打痕」という単独項目が無く、チャック/回収/落下/材料/加工時/ガイドブッシュ/流動の
// どの工程で付いたキズかで項目が分かれているため、まず発生工程の語を判定し、特定できなければ
// 「その他キズ」に丸める。「材質違い」は「材料キズ」と語が紛らわしいため先に判定する。
var KP_KEYWORD_RULES_ = [
  [/材質.*違い|材料.*違い/, '材質違い'],
  [/挽目|面粗度/, '挽目不良'],
  [/バフ/, 'バフがけ不良'],
  [/取り残し/, '取り残し'],
  [/ムシレ|むしれ/, 'ムシレ'],
  [/切粉/, '切粉'],
  [/汚れ/, '汚れ'],
  [/変色|錆/, '変色・錆'],
  [/段差/, '段差'],
  [/溶着/, '溶着'],
  [/バリ/, 'バリ'],
  [/修正不良/, '修正不良'],
  [/チャック/, 'チャックキズ'],
  [/回収/, '回収時キズ'],
  [/落下/, '落下キズ'],
  [/材料/, '材料キズ'],
  [/加工時|加工中/, '加工時キズ'],
  [/ガイドブッシュ/, 'ガイドブッシュキズ'],
  [/流動|洗浄|移し替え/, '流動キズ'],
  [/圧痕|打痕|キズ|傷|かじり|カジリ/, 'その他キズ'],
  [/(ねじ|ネジ).*(不良|曲がり|通り|抜け)/, 'ねじ不良'],
  [/穴ズレ|穴曲/, '穴ズレ（穴曲り）'],
  [/振れ/, '振れ大'],
  [/テーパ/, 'テーパ不良'],
  [/平研/, '平研不良'],
  [/(長さ|全長).*(不良|短|長)/, '長さ不良'],
  [/外径.*(大|プラス|\+)/, '外径大'],
  [/外径.*(小|マイナス|\-)/, '外径小'],
  [/内径.*(大|プラス|\+)/, '内径大'],
  [/内径.*(小|マイナス|\-)/, '内径小'],
  [/未加工/, '未加工混入（未加工）'],
  [/圧入|接着/, '圧入・接着不良'],
  [/現品/, '現品違い'],
  [/曲がり/, '曲がり'],
  [/変形/, '変形'],
  [/2次加工/, '寸法不良']
];

/**
 * 旧項目名(自由記述)を新マスタ(DEFECT_ITEMS、40項目)の名称にできるだけ近く丸める。
 * ①完全一致 → ②表記ゆれのエイリアス → ③項目名自体のキーワード判定 → ④キズ原因タグ(あれば)のキーワード判定
 * → ⑤それでも不明なら「形状不良」(新マスタに万能の「その他」項目が無いため、最も汎用的な項目に丸める)。
 * 元のテキストは呼び出し側で必ず不良項目詳細列に残すため、丸めても情報は失われない。
 */
function mapDefectItemName_(oldName, scratchTag, log) {
  var itemNames = DEFECT_ITEMS.map(function (item) { return item.name; });
  if (itemNames.indexOf(oldName) !== -1) return oldName;

  var aliasMap = {
    '外径G/B傷': 'その他キズ',
    '外径ｇ/ｂ傷': 'その他キズ',
    '寸法出し': '寸法不良'
  };
  if (aliasMap[oldName]) return aliasMap[oldName];

  for (var i = 0; i < KP_KEYWORD_RULES_.length; i++) {
    if (KP_KEYWORD_RULES_[i][0].test(oldName)) {
      log.push('  ※旧項目「' + oldName + '」→「' + KP_KEYWORD_RULES_[i][1] + '」に判定しました(元の名称は不良項目詳細に残しています)');
      return KP_KEYWORD_RULES_[i][1];
    }
  }

  // 項目名自体からは判定できない場合、旧Googleフォームの「キズ原因」タグ(あれば)から補助的に判定する
  if (scratchTag) {
    var tag = scratchTag.toString();
    for (var j = 0; j < KP_KEYWORD_RULES_.length; j++) {
      if (KP_KEYWORD_RULES_[j][0].test(tag)) {
        log.push('  ※旧項目「' + oldName + '」→キズ原因タグ「' + tag + '」から「' + KP_KEYWORD_RULES_[j][1] + '」に判定しました(元の名称は不良項目詳細に残しています)');
        return KP_KEYWORD_RULES_[j][1];
      }
    }
  }

  log.push('  ※旧項目「' + oldName + '」が新マスタのどれにも判定できなかったため「形状不良」に丸めました(元の名称は不良項目詳細に残しています)');
  return '形状不良';
}

/**
 * 旧Googleフォームの「キズ原因」選択値(例:「B.チャック内の切粉による圧痕」)を、
 * 新マスタ(KP_CAUSE_ITEMS)の名称(例:「チャック内の切粉による圧痕」)に変換する。
 * プレフィックス(コード+ドット)を外して完全一致するものだけ採用し、一致しなければ空欄を返す(無理に丸めない)。
 */
function mapScratchToKizuGenin_(scratchTag) {
  if (!scratchTag) return '';
  var stripped = scratchTag.toString().replace(/^[A-Zａ-ｚA-Ｚ][\.．]\s*/, '').trim();
  var causeNames = KP_CAUSE_ITEMS.map(function (item) { return item.name; });
  return causeNames.indexOf(stripped) !== -1 ? stripped : '';
}

/** "26/06/02" のような文字列・Dateオブジェクトのどちらでも日付として読み取る */
function parseSourceDate_(value) {
  if (value instanceof Date) return value;
  var s = value ? value.toString().trim() : '';
  var m = s.match(/(\d{2,4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    var y = Number(m[1]);
    if (y < 100) y += 2000;
    return new Date(y, Number(m[2]) - 1, Number(m[3]));
  }
  return new Date();
}

/** レコード配列を、指定シートの次の空き行から書き込む(WebApi.gsのwriteDefectRecord_と同じ列構成) */
function writeRecordsToSheet_(sheet, records) {
  var writtenRows = 0;
  records.forEach(function (rec) {
    var row = findNextEmptyRow_(sheet);
    if (row > MIGRATE_DATA_END_ROW) {
      Logger.log('【エラー】' + sheet.getName() + ' が満杯のため、これ以上書き込めません。');
      return;
    }
    var totalQty = rec.items.reduce(function (sum, it) { return sum + (Number(it.qty) || 0); }, 0);

    sheet.getRange(row, 1, 1, 9).setValues([[
      rec.timestamp, rec.shochiKubun, rec.quality, rec.customer, rec.drawing,
      rec.worker, rec.kishu, rec.setsubi, Number(rec.suryo) || ''
    ]]);
    sheet.getRange(row, 11, 1, 5).setValues([[
      totalQty, rec.items[0].name, Number(rec.items[0].qty) || '', rec.items[0].detail || '', rec.items[0].worker2 || ''
    ]]);
    sheet.getRange(row, 16).setValue(Number(rec.price) || ''); // P 単価
    sheet.getRange(row, 18).setValue(rec.remark || '');        // R 備考
    sheet.getRange(row, 19).setValue(rec.material || '');      // S 材質
    sheet.getRange(row, 21).setValue(rec.kizuGenin || '');     // U キズ原因
    sheet.getRange(row, 22).setValue(MIGRATE_MARKER);          // V 送信ID(移行行の目印)
    writtenRows++;

    for (var i = 1; i < rec.items.length; i++) {
      var r2 = findNextEmptyRow_(sheet);
      if (r2 > MIGRATE_DATA_END_ROW) {
        Logger.log('【エラー】' + sheet.getName() + ' が満杯のため、不良項目の追加行を書き込めません。');
        break;
      }
      sheet.getRange(r2, 2).setValue(rec.shochiKubun);
      sheet.getRange(r2, 12, 1, 4).setValues([[rec.items[i].name, Number(rec.items[i].qty) || '', rec.items[i].detail || '', rec.items[i].worker2 || '']]);
      sheet.getRange(r2, 22).setValue(MIGRATE_MARKER); // V 送信ID(移行行の目印、undoJuneJulyMigrationで使う)
      writtenRows++;
    }
  });
  return writtenRows;
}

/** A列(タイムスタンプ)・L列(不良項目)が両方空の最初の行を返す(WebApi.gsのfindNextRow_と同じ考え方) */
function findNextEmptyRow_(sheet) {
  var values = sheet.getRange(2, 1, MIGRATE_DATA_END_ROW - 1, 12).getValues();
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0] && !values[i][11]) return i + 2;
  }
  return MIGRATE_DATA_END_ROW + 1;
}
