// カスタムメニューを作成し、メニュー項目を追加する
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('新規スペース登録', 'addNewRecord')
    .addItem('PJCD変更', 'replaceValues')
    .addItem('スペース削除', 'deleteSpace')
    .addToUi();
}

// 値を入力するためのプロンプトを表示し、入力された値を返す
function showPrompt(message) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() === ui.Button.OK) {
    return result.getResponseText();
  } else {
    return null;
  }
}

// メッセージを表示するアラートを表示する
function showMessage(message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

/////////////////////////////////////////////////////////////////////////////

// カスタムメニュー「新規スペース登録」が選択された時に実行される関数
function addNewRecord() {
  // 関数実行先のシートを設定する
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  sheet.activate();

  // B列の値を入力するためのプロンプトを表示し、入力を受け取る
  var newValueB = showPrompt('新しいKey(B列)の値を入力してください');
  if (newValueB === null) return; // キャンセルされた場合は処理を終了

  // B列に同じKeyの値が存在するか確認
  var valuesB = sheet.getRange('B:B').getValues();
  for (var i = 0; i < valuesB.length; i++) {
    if (valuesB[i][0] === newValueB) {
      showMessage('B列に同じ値が既に存在します。別の値を入力してください。');
      newValueB = showPrompt('新しいKey(B列)の値を再度入力してください');
      if (newValueB === null) return; // キャンセルされた場合は処理を終了
      i = -1; // ループを再度実行するためにインデックスをリセット
    }
  }

  // enable_disable(A列)に1を自動的に入力
  var newValueA = 1;

  // C列からG列の値を入力するためのプロンプトを表示し、入力を受け取る
  var newValueC = showPrompt('新しいSpaceName(C列)の値を入力してください');
  if (newValueC === null) return;

  var newValueD = showPrompt('新しいPJCD(D列)の値を入力してください');
  if (newValueD === null) return;

  var newValueE = showPrompt('新しいPJCDName(E列)の値を入力してください');
  if (newValueE === null) return;

  var newValueF = showPrompt('新しいOverview(F列)の値を入力してください');
  if (newValueF === null) return;

  var newValueG = showPrompt('新しいContractNumber(G列)の値を入力してください');
  if (newValueG === null) return;

  // 新しいレコードの値を設定
  var newRow = [newValueA, newValueB, newValueC, newValueD, newValueE, newValueF, newValueG];

  // シートの最終行に新しいレコードを追加
  sheet.appendRow(newRow);

  // 追加した新しいレコードにジャンプする
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).activate();

  // 追加結果を表示する
  showMessage('新しいレコードを追加しました');
}

/////////////////////////////////////////////////////////////////////////////

// カスタムメニュー「PJCD変更」が選択された時に実行される関数
// ユーザーから3つの値を入力し、それらの値を使って処理を実行する
function replaceValues() {
  // 関数実行先のシートを設定する
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  sheet.activate();

  // 検索する値を入力するプロンプトを表示し、入力を受け取る
  var searchValue = showPrompt('変更対象のKey(B列)を入力してください');
  if (searchValue === null) return; // キャンセルされた場合は処理を終了

  // 置換する値を入力するプロンプトを表示し、入力を受け取る
  var replaceValue1 = showPrompt('置換するPJCD(D列)を入力してください');
  if (replaceValue1 === null) return; // キャンセルされた場合は処理を終了

  // 置換する値をもう1つ入力するプロンプトを表示し、入力を受け取る
  var replaceValue2 = showPrompt('置換するPJCDName(E列)を入力してください');
  if (replaceValue2 === null) return; // キャンセルされた場合は処理を終了

  // Bカラムの値を配列として取得する
  var range = sheet.getRange('B:B');
  var values = range.getValues();
  var found = false; // 値が見つかったかどうかを表すフラグ

  // Bカラム内で検索する
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === searchValue) {
      var row = i + 1; // 値が見つかった行
      // 同行のDカラムの値を置換する
      sheet.getRange(row, 4).setValue(replaceValue1);
      // 同行のEカラムの値を置換する
      sheet.getRange(row, 5).setValue(replaceValue2);
      // 置換が行われたセルにジャンプする
      sheet.getRange(row, 1).activate();
      found = true; // 値が見つかったことを示す
    }
  }

  // 置換結果を表示する
  if (found) {
    showMessage('置換が成功しました');
  } else {
    showMessage('入力したKeyは存在しませんでした');
  }
}

/////////////////////////////////////////////////////////////////////////////

// カスタムメニュー「スペース削除」が選択された時に実行される関数
function deleteSpace() {
  // 関数実行先のシートを設定する
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  sheet.activate();

  // B列のKeyの入力を受け取る
  var searchValue = showPrompt('削除対象スペースのKey(B列)を入力してください');
  if (searchValue === null) return; // キャンセルされた場合は処理を終了

  // B列に同じKeyの値が存在するか確認
  var valuesB = sheet.getRange('B:B').getValues();
  var found = false; // 値が見つかったかどうかを表すフラグ

  for (var i = 0; i < valuesB.length; i++) {
    if (valuesB[i][0] === searchValue) {
      found = true; // 値が見つかったことを示す
      var row = i + 1; // 値が見つかった行
      var valueA = sheet.getRange(row, 1).getValue(); // A列の値を取得

      if (valueA === 1) {
        sheet.getRange(row, 1).setValue(0);
        sheet.getRange(row, 1).activate(); // 置換が行われたセルにジャンプ
        showMessage('該当のスペースを削除しました');
      } else if (valueA === 0) {
        sheet.getRange(row, 1).activate(); // 置換が行われたセルにジャンプ
        showMessage('該当のスペースは既に削除されています');
      }

      break; // 一致するKeyが見つかったらループを終了
    }
  }

  // 一致するKeyが見つからない場合のメッセージを表示
  if (!found) {
    showMessage('一致するKeyは存在しませんでした');
  }
}