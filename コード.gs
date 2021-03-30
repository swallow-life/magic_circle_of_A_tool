const BASE_ROW = 2;
const MAX_SUCCESS_ELEMENT = 16;
const DOC_PROP_SHEET_NUMBER = "sheetNumber";
const CONFIG_SHEET_NAME = "設定";

const number_half_wide_map = {
  0: "０",
  1: "１",
  2: "２",
  3: "３",
  4: "４",
  5: "５",
  6: "６",
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('成功要素管理')
      .addItem('成功要素成長', 'create_result')
      .addSeparator()
      .addItem('新規成功要素登録', 'add_new_success_element')
      .addItem('成功要素テキスト表示', 'show_success_element')
      .addSeparator()
      .addItem('使い方', 'help')
      // .addItem('シート番号リセット', 'reset_sheet_number')
      .addToUi();
}

function debug() {
  let v = [1,2,3].length;
  console.log(v);
  // let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1");
  // let same_time = sheet.getRange(1, 7).getValue();
  // console.log(same_time);
}

function reset_sheet_number() {
  let documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty(DOC_PROP_SHEET_NUMBER);
}

function help() {
  let html = HtmlService.createHtmlOutputFromFile('help');
  html.setWidth(850);
  html.setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '簡単な使い方');
}

function create_result() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }

  console.log(sheet.getSheetName());
  // "A2:D17"
  let a1notation = "A" + BASE_ROW +":D" + (BASE_ROW + MAX_SUCCESS_ELEMENT - 1);
  let range = sheet.getRange(a1notation);

  let same_time = sheet.getRange(1, 6).getValue();
  console.log(same_time);

  // シートから値を取得する
  let vals = [];
  for (let i = 1; i <= MAX_SUCCESS_ELEMENT; i++) {
    let target = range.getCell(i, 1).getValue();
    let name = range.getCell(i, 2).getDisplayValue();
    if (name === "") {
      continue;
    }
    let power = range.getCell(i, 3).getValue();
    let count = range.getCell(i, 4).getValue();
    vals.push({target: target, name: name, power: power, count: count, prev: "", available: true})
  }

  let numSuccessElement = vals.filter(x => x.available === true).length;
  let results = vals.flatMap(val => {
    if (val.target !== true) {
      // 使わなかった成功要素（連続使用回数を0にする）
      return [{name: val.name, power: val.power, count: 0, prev: "", available: true}];
    }

    let name_power = val.name + "(" + val.power + ")";
    let nextPower = Math.min(6, val.power + 1);
    let nextCount;
    if (val.count + 1 === 3) {
      let tempArr = [];
      let available1 = true;
      let available2 = true;
      // 分割
      nextCount = 0;
      nextPower--;
      let nextName1
      if (numSuccessElement < MAX_SUCCESS_ELEMENT) {
        nextName1 = inputBoxCustum_(name_power + "の成長分割1");
        numSuccessElement++;
      } else {
        nextName1 = "（分割時、最大成功要素数を超過のため削除）";
        available1 = false;
      }
      tempArr.push({
        name: nextName1,
        power: nextPower,
        count: nextCount,
        prev: name_power + "からの成長分割",
        available: available1
      });
      let nextName2
      if (numSuccessElement < MAX_SUCCESS_ELEMENT) {
        nextName2 = inputBoxCustum_(name_power + "の成長分割2");
        numSuccessElement++;
      } else {
        nextName2 = "（分割時、最大成功要素数を超過のため削除）";
        available2 = false;
      }
      tempArr.push({
        name: nextName2,
        power: nextPower,
        count: nextCount,
        prev: name_power + "からの成長分割",
        available: available2
      });
      numSuccessElement--;
      return tempArr;
    } else {
      // 成長
      nextCount = val.count + 1;
      // 同時提出の場合はリセットする
      if (same_time === true) {
        nextCount = 0;
      }
      let nextName;
      let prev;
      if (val.power < 6) {
        nextName = inputBoxCustum_(name_power + "の成長");
        prev = name_power + "からの成長";
      } else {
        // パワー６の場合名前は変わらない
        nextName = val.name;
        prev = "";
      }
      return [{name: nextName, power: nextPower, count: nextCount, prev: prev, available: true}];
    }
  });
  let template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("template");
  let copySheet = template.copyTo(SpreadsheetApp.getActiveSpreadsheet());

  let documentProperties = PropertiesService.getDocumentProperties();  
  console.log(documentProperties.getKeys().some((value) => value === DOC_PROP_SHEET_NUMBER));
  let sheetNumber = documentProperties.getKeys().some((value) => value === DOC_PROP_SHEET_NUMBER)
    ? documentProperties.getProperty(DOC_PROP_SHEET_NUMBER)
    : 1; // 初回起動時
  sheetNumber++;
  documentProperties.setProperty(DOC_PROP_SHEET_NUMBER, sheetNumber);

  copySheet.setName(sheetNumber);
  copySheet.activate();
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(2);
  let dataRange = copySheet.getDataRange();
  let index = BASE_ROW;
  for (let re in results) {
    if (results[re].available === false) {
      continue;
    }
    dataRange.getCell(index, 2).setValue(results[re].name);
    dataRange.getCell(index, 3).setValue(results[re].power);
    dataRange.getCell(index, 4).setValue(results[re].count);
    if (results[re].prev !== "") {
      dataRange.getCell(index, 2).setNote(results[re].prev);
    }
    index++;
  }
  let inAvailableRange = copySheet.getRange("A18:D33");
  let inAvailableIndex = 1;
  for (let re in results) {
    if (results[re].available === true) {
      continue;
    }
    inAvailableRange.getCell(inAvailableIndex, 2).setValue(results[re].name);
    inAvailableRange.getCell(inAvailableIndex, 3).setValue(results[re].power);
    inAvailableRange.getCell(inAvailableIndex, 4).setValue(results[re].count);
    if (results[re].prev !== "") {
      inAvailableRange.getCell(inAvailableIndex, 2).setNote(results[re].prev);
    }
    inAvailableIndex++;
  }

  let today = new Date();
  // アメリカ東海岸時間-4から日本時間+9に変換するので+13
  today.setHours(today.getHours() + 13);
  copySheet.getRange(1, 7).setValue(today.toLocaleString());
}

function inputBoxCustum_(initMessage) {
  let inputValue = Browser.inputBox(initMessage);
  if (inputValue === 'cancel') {
    inputValue = "";
  }
  return inputValue;
}

function add_new_success_element() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }

  let dataRange = sheet.getDataRange();
  for (let i = BASE_ROW; i < BASE_ROW + MAX_SUCCESS_ELEMENT - 1; i++) {
    let name = sheet.getRange('B' + i).getDisplayValue();
    if (name === "") {
      let newName = inputBoxCustum_("新規成功要素の登録");
      dataRange.getCell(i, 2).setValue(newName);
      dataRange.getCell(i, 3).setValue(2);
      dataRange.getCell(i, 4).setValue(0);
      dataRange.getCell(i, 2).setNote("新規成功要素");
      return;
    }
  }
  Browser.msgBox("成功要素の数が最大のため新規成功要素は登録できません");
}

function show_success_element() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }
  console.log(sheet.getSheetName());

  let config_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  let config_range = config_sheet.getRange("A1:I6");

  let html = HtmlService.createTemplateFromFile('index');
  
  let results = [];
  // a1notation is like "A2:D17"
  let a1notation = "A" + BASE_ROW +":D" + (BASE_ROW + MAX_SUCCESS_ELEMENT - 1);
  let range = sheet.getRange(a1notation);
  for (let i = 1; i <= MAX_SUCCESS_ELEMENT; i++) {
    let name = range.getCell(i, 2).getDisplayValue();
    if (name === "") {
      continue;
    }
    let power = range.getCell(i, 3).getValue();
    let count = range.getCell(i, 4).getValue();
    results.push(custom_format_(config_range, name, power, count));    
  }

  html.data = results;
  SpreadsheetApp.getUi().showModalDialog(html.evaluate(), '成功要素のテキスト表示');
}

function custom_format_(config_range, name, power, count) {
  let pre = config_range.getCell(4, 1).getValue();
  let pre_power = config_range.getCell(4, 3).getValue();
  let post_power = config_range.getCell(4, 5).getValue();
  let pre_count = config_range.getCell(4, 6).getValue();
  let post_count = config_range.getCell(4, 8).getValue();
  let post = config_range.getCell(4, 9).getValue();
  let symbol_as_count = config_range.getCell(6, 1).getValue();
  let count_str;
  if (symbol_as_count === '') {
    count_str = count === 0 ? "" : `${pre_count}${number_half_wide_map[count]}${post_count}`;
  } else {
    count_str = count === 0 ? "" : `${pre_count}${symbol_as_count.repeat(count)}${post_count}`;
  }
  return `${pre}${name}${pre_power}${number_half_wide_map[power]}${post_power}${count_str}${post}`;
}
