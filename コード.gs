const BASE_ROW = 4;
const MAX_SUCCESS_ELEMENT = 16;
const LAST_AVAILABLE_ROW = BASE_ROW + MAX_SUCCESS_ELEMENT - 1;
const LAST_INAVAILABLE_ROW = LAST_AVAILABLE_ROW + 1 + MAX_SUCCESS_ELEMENT - 1;

const DOC_PROP_SHEET_NUMBER = "sheetNumber";
const CONFIG_SHEET_NAME = "設定";
const NEW_SUCCESS_ELEMENT = "新規成功要素";
const MAX_SUCCESS_ELEMENT_EXCEED = "（分割時、最大成功要素数を超過のため削除）";
const INAVAILABLE_COLOR = "grey";

const MAX_GENKAI_TOPPA_POWER = 20;
const MAX_POWER = 6;

const number_half_wide_map = {
  0:  "０",
  1:  "１",
  2:  "２",
  3:  "３",
  4:  "４",
  5:  "５",
  6:  "６",
  7:  "７",
  8:  "８",
  9:  "９",
  10: "１０",
  11: "１１",
  12: "１２",
  13: "１３",
  14: "１４",
  15: "１５",
  16: "１６",
  17: "１７",
  18: "１８",
  19: "１９",
  20: "２０",
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('成功要素管理')
      .addItem('成功要素成長', 'create_result')
      .addSeparator()
      .addItem('新規成功要素登録', 'add_new_success_element')
      .addItem('成功要素停止', 'stop_success_element')
      .addSeparator()
      .addItem('キャラクターシート用テキスト表示', 'show_success_element')
      .addItem('統制判定提出用テキスト表示', 'show_target_success_element')
      .addItem('成長申請用テキスト表示', 'show_result')
      .addSeparator()
      .addItem('対象チェックボックス全チェック', 'target_all_check')
      .addItem('対象チェックボックスリセット', 'target_reset')
      .addSeparator()
      .addItem('使い方', 'help')
      .addToUi();

  let documentProperties = PropertiesService.getDocumentProperties();  
  if (documentProperties.getKeys().some(key => key === DOC_PROP_SHEET_NUMBER) === false) {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1").activate();
  }

}

function debug() {
  const preserve_sheet_names = ["設定", "template"];
  let activeSpSh = SpreadsheetApp.getActiveSpreadsheet();
  activeSpSh.getSheets()
    .filter(x => !preserve_sheet_names.includes(x.getSheetName()))
    .forEach(x => activeSpSh.deleteSheet(x))
    ;
}

function init() {
  delete_data_sheets();
  create_first_sheet();
  reset_sheet_number();
}

function delete_data_sheets() {
  const preserve_sheet_names = ["設定", "template"];
  let active = SpreadsheetApp.getActiveSpreadsheet();
  active.getSheets()
    .filter(x => !preserve_sheet_names.includes(x.getSheetName()))
    .forEach(x => active.deleteSheet(x))
  ;
}

function create_first_sheet() {
  let template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("template");
  let copySheet = template.copyTo(SpreadsheetApp.getActiveSpreadsheet());
  copySheet.setName("1");
  copySheet.activate();
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(2);
  let range = copySheet.getDataRange();
  range.getCell(4, 2).setValue("（初期成功要素を入力してください）");
  range.getCell(4, 3).setValue(6);
  range.getCell(4, 4).setValue(0);
  range.getCell(5, 2).setValue("（初期成功要素を入力してください）");
  range.getCell(5, 3).setValue(6);
  range.getCell(5, 4).setValue(0);
  reset_sheet_number();
}

function reset_sheet_number() {
  let documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty(DOC_PROP_SHEET_NUMBER);
}

function help() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(HtmlService.createHtmlOutputFromFile('help')
      .setTitle('簡単な使い方'));
}

function target_all_check() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }

  let range = get_data_range_(sheet);
  for (let row = 1; row <= MAX_SUCCESS_ELEMENT; row++) {
    let name = range.getCell(row, 2).getDisplayValue();
    if (name === "") {
      continue;
    }
    range.getCell(row, 1).setValue(true);
  }
}

function target_reset() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }

  let range = get_data_range_(sheet);
  for (let row = 1; row <= MAX_SUCCESS_ELEMENT; row++) {
    range.getCell(row, 1).setValue(false);
  }
}

function create_result() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }

  // 設定シートから各種設定値を取得する
  let config_range = get_config_range_();
  let config = read_config_(config_range);
  let use_dialog = config_range.getCell(12, 1).getValue();

  let douji_kadai = sheet.getRange(1, 1).getValue();
  let genkai_toppa = sheet.getRange(2, 1).getValue();
  let kenshoku = config_range.getCell(14, 1).getValue();

  let sheet_name = sheet.getSheetName();
  console.log(sheet_name);
  // "A4:D19"
  let range = get_data_range_(sheet);

  // シートから成功要素の各値を取得する
  let vals = [];
  for (let row = 1; row <= MAX_SUCCESS_ELEMENT; row++) {
    let target = range.getCell(row, 1).getValue();
    let name = range.getCell(row, 2).getDisplayValue();
    if (name === "") {
      continue;
    }
    let power = range.getCell(row, 3).getValue();
    let power_count = "";
    if (kenshoku === true) {
      power_count = range.getCell(row, 3).getNote();
    }
    let count = range.getCell(row, 4).getValue();
    vals.push({target: target, name: name, power: power, count: count, note: "", available: true, power_count: power_count})
  }

  let numSuccessElement = vals.filter(x => x.available === true).length;

  let results = vals.flatMap(val => {
    if (val.target !== true) {
      // 使わなかった成功要素（連続使用回数を0にする）
      return [{name: val.name, power: val.power, count: 0, note: "", available: true, power_count: val.power_count}];
    }

    const name_power = custom_format_(config, val.name, val.power, val.count);
    let nextPower;
    let nextPowerCount = "";
    let nextCount;
    let note;

    // 限界突破の成長処理（分割ルールを停止）
    if (genkai_toppa === true && numSuccessElement === MAX_SUCCESS_ELEMENT) {
      nextCount = val.count + 1;
      if (douji_kadai === true) {
        // 同時提出の場合は連続使用回数をリセットする
        nextCount = 0;
      }

      let nextName;
      if (val.power >= MAX_POWER) {
        // パワー６以上の場合の処理
        nextName = val.name;
        if (val.count + 1 === 20) {
          // 二〇回連続で使用された成功要素はパワーを＋１する
          nextPower = power_up_(MAX_GENKAI_TOPPA_POWER, val.power);
          nextCount = 0;
          note = name_power + "からの成長（限界突破）";
        } else {
          // パワーそのままで連続使用回数のカウントアップのみ
          nextPower = val.power;
          note = "";
        }
      } else {
        // パワー５以下の場合の処理
        if (kenshoku === false || val.power_count === "1") {
          // 兼職していない場合と兼職で二回目の使用の場合は成長処理を実施する
          nextName = use_dialog ? inputBoxCustum_(name_power + "の成長") : '';
          nextPower = power_up_(MAX_POWER, val.power);
          note = name_power + "からの成長";
        } else {
          // 兼職で一回目の使用の場合、連続使用回数をカウントアップし、さらにパワーの回数をカウントアップ
          nextName = val.name;
          nextPower = val.power;
          nextPowerCount = "1";
          note = "";
        }
      }

      return [{name: nextName, power: nextPower, count: nextCount, note: note, available: true, power_count: nextPowerCount}];
    }

    // 限界突破以外の成長処理
    let maxCount;
    if (kenshoku === true) {
      // 兼職の処理
      if (val.power_count === "1") {
        // 二回目の使用でパワーが上がる
        nextPower = power_up_(MAX_POWER, val.power);
        nextPowerCount = "";
      } else {
        // 一回目の使用ではパワーそのまま
        nextPower = val.power;
        if (val.power < MAX_POWER) {
          nextPowerCount = "1";
        }
      }
      maxCount = 6;
    } else {
      // 通常の（兼職でない）処理
      nextPower = power_up_(MAX_POWER, val.power);
      maxCount = 3;
    }
    if (val.count + 1 >= maxCount) {
      // 分割処理
      let dividedArr = [];
      let available1 = true;
      let available2 = true;

      nextCount = 0;
      nextPower--;
      numSuccessElement--;

      let nextName1
      if (numSuccessElement < MAX_SUCCESS_ELEMENT) {
        nextName1 = use_dialog ? inputBoxCustum_(name_power + "の成長分割1") : '';
        numSuccessElement++;
      } else {
        nextName1 = MAX_SUCCESS_ELEMENT_EXCEED;
        available1 = false;
      }
      dividedArr.push({
        name: nextName1, power: nextPower, count: nextCount, note: name_power + "からの成長分割", available: available1, power_count: "",
      });

      let nextName2
      if (numSuccessElement < MAX_SUCCESS_ELEMENT) {
        nextName2 = use_dialog ? inputBoxCustum_(name_power + "の成長分割2") : '';
        numSuccessElement++;
      } else {
        nextName2 = MAX_SUCCESS_ELEMENT_EXCEED;
        available2 = false;
      }
      dividedArr.push({
        name: nextName2, power: nextPower, count: nextCount, note: name_power + "からの成長分割", available: available2, power_count: "",
      });

      return dividedArr;
    } else {
      // 成長処理
      nextCount = val.count + 1;
      if (douji_kadai === true) {
        // 同時提出の場合は連続使用回数をリセットする
        nextCount = 0;
      }
      let nextName;
      if (val.power < MAX_POWER && (kenshoku === false || val.power_count === "1")) {
        nextName = use_dialog ? inputBoxCustum_(name_power + "の成長") : '';
        note = name_power + "からの成長";
      } else {
        // パワー６の場合、もしくは兼職で一回目の使用の場合、名前は変わらない
        nextName = val.name;
        note = "";
      }
      return [{name: nextName, power: nextPower, count: nextCount, note: note, available: true, power_count: nextPowerCount}];
    }
  });

  let template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("template");
  let copySheet = template.copyTo(SpreadsheetApp.getActiveSpreadsheet());

  // シートの名前を作る（連番）
  let documentProperties = PropertiesService.getDocumentProperties();  
  console.log(documentProperties.getKeys().some((value) => value === DOC_PROP_SHEET_NUMBER));
  let sheetNumber = documentProperties.getKeys().some((value) => value === DOC_PROP_SHEET_NUMBER)
    ? documentProperties.getProperty(DOC_PROP_SHEET_NUMBER)
    : 1; // 初回起動時
  sheetNumber++;
  documentProperties.setProperty(DOC_PROP_SHEET_NUMBER, sheetNumber);
  copySheet.setName(sheetNumber);

  let dataRange = copySheet.getDataRange();
  results
    .filter(x => x.available === true)
    .forEach((result, i) => {
      let row = i + BASE_ROW;
      let nameCell = dataRange.getCell(row, 2);
      nameCell.setValue(result.name);
      if (result.note !== "") {
        nameCell.setNote(result.note);
        nameCell.setBackground("lightblue");
      }
      dataRange.getCell(row, 3).setValue(result.power);
      if (result.power_count != null && result.power_count === "1") {
        dataRange.getCell(row, 3).setNote(result.power_count);
      }
      dataRange.getCell(row, 4).setValue(result.count);
    });
  // 成功要素を分割した際に最大数を超えた分を扱う
  let inAvailableA1notation = "A" + (LAST_AVAILABLE_ROW + 1) +":D" + LAST_INAVAILABLE_ROW;
  let inAvailableRange = copySheet.getRange(inAvailableA1notation);
  results
    .filter(x => x.available === false)
    .forEach((result, i) => {
      let inAvailableRow = i + 1;
      let nameCell = inAvailableRange.getCell(inAvailableRow, 2);
      let powerCell = inAvailableRange.getCell(inAvailableRow, 3);
      let countCell = inAvailableRange.getCell(inAvailableRow, 4);

      nameCell.setValue(result.name);
      powerCell.setValue(result.power);
      countCell.setValue(result.count);
      nameCell.setBackground(INAVAILABLE_COLOR);
      powerCell.setBackground(INAVAILABLE_COLOR);
      countCell.setBackground(INAVAILABLE_COLOR);

      if (result.note !== "") {
        nameCell.setNote(result.note);
      }
    });
  // 左の端は設定シートにしたいので左から2番目に成長結果のシートを移動させる
  copySheet.activate();
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(2);

  copySheet.getRange(2, 1).setValue(genkai_toppa);
  let today = new Date();
  // アメリカ東海岸時間-4から日本時間+9に変換するので+13
  today.setHours(today.getHours() + 13);
  copySheet.getRange(1, 4).setValue(today.toLocaleString());
  copySheet.getRange(2, 4).setValue(sheet_name);
}

function add_new_success_element() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }

  // let config_range = get_config_range_();
  let use_dialog = true;//config_range.getCell(12, 1).getValue();

  let dataRange = sheet.getDataRange();
  for (let row = BASE_ROW; row <= LAST_AVAILABLE_ROW; row++) {
    let name = dataRange.getCell(row, 2).getDisplayValue();
    let note = dataRange.getCell(row, 2).getNote();
    if (note === NEW_SUCCESS_ELEMENT) {
      Browser.msgBox("すでに新規成功要素は登録されています");
      return;
    }
    if (name === "" && note === "") {
      let newName = use_dialog ? inputBoxCustum_("新規成功要素の登録") : '';
      dataRange.getCell(row, 2).setValue(newName);
      dataRange.getCell(row, 3).setValue(2);
      dataRange.getCell(row, 4).setValue(0);
      dataRange.getCell(row, 2).setNote(NEW_SUCCESS_ELEMENT);
      return;
    }
  }
  Browser.msgBox("成功要素の数が最大のため新規成功要素は登録できません");
}

function stop_success_element() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }

  // "A4:D19"
  let range = get_data_range_(sheet);

  let stopArr = [];
  let selected = false;
  for (let row = 1; row <= MAX_SUCCESS_ELEMENT; row++) {
    let target = range.getCell(row, 1).getValue();
    if (target === false) {
      continue;
    }
    selected = true;
    range.getCell(row, 1).setValue(false);

    let nameCell = range.getCell(row, 2);
    let name = nameCell.getDisplayValue();

    let powerCell = range.getCell(row, 3);
    let power = powerCell.getValue();

    let powerCount = powerCell.getNote();

    let countCell = range.getCell(row, 4);
    let count = countCell.getValue();

    let note = nameCell.getNote();
    nameCell.clearNote();
    powerCell.clearContent();
    powerCell.clearNote();
    countCell.clearContent();
    nameCell.clearContent();

    if (note === NEW_SUCCESS_ELEMENT) {
      continue;
    }
    stopArr.push({target: target, name: name, power: power, count: count, note: note, available: true, power_count: powerCount})
  }
  if (selected === false) {
    Browser.msgBox("停止したい成功要素を選択してください");
    return;
  }

  let inAvailableA1notation = "A" + (LAST_AVAILABLE_ROW + 1) +":D" + LAST_INAVAILABLE_ROW;
  let inAvailableRange = sheet.getRange(inAvailableA1notation);
  let inAvailableRow = 1;
  stopArr.forEach(stop =>  {
    for (; inAvailableRow <= MAX_SUCCESS_ELEMENT;) {
      let nameCell = inAvailableRange.getCell(inAvailableRow, 2);
      if (nameCell.getDisplayValue() !== "") {
        inAvailableRow++;
        continue;
      }
      let powerCell = inAvailableRange.getCell(inAvailableRow, 3);
      let countCell = inAvailableRange.getCell(inAvailableRow, 4);

      nameCell.setValue(stop.name);
      powerCell.setValue(stop.power);
      if (stop.power_count != null && stop.power_count === "1") {
        inAvailableRange.getCell(inAvailableRow, 3).setNote(stop.power_count);
      }

      countCell.setValue(stop.count);
      nameCell.setBackground(INAVAILABLE_COLOR);
      powerCell.setBackground(INAVAILABLE_COLOR);
      countCell.setBackground(INAVAILABLE_COLOR);
      if (stop.note !== "") {
        nameCell.setNote(stop.note);
      }
      inAvailableRow++;
      break;
    }
  });

}

function show_success_element() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }
  console.log(sheet.getSheetName());

  let config_range = get_config_range_();
  let config = read_config_(config_range);

  let results = [];
  // "A2:D17"
  let range = get_data_range_(sheet);
  for (let row = 1; row <= MAX_SUCCESS_ELEMENT; row++) {
    let name = range.getCell(row, 2).getDisplayValue();
    if (name === "") {
      continue;
    }
    let power = range.getCell(row, 3).getValue();
    let count = range.getCell(row, 4).getValue();
    results.push(custom_format_(config, name, power, count));    
  }

  let html = HtmlService.createTemplateFromFile('index');
  html.data = results;
  let evalHtml = html.evaluate();
  evalHtml.setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(evalHtml, '成功要素のテキスト表示');
}

function show_target_success_element() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }
  console.log(sheet.getSheetName());

  let config_range = get_config_range_();
  let config = read_config_(config_range);

  let results = [];
  // "A2:D17"
  let range = get_data_range_(sheet);
  for (let row = 1; row <= MAX_SUCCESS_ELEMENT; row++) {
    let target = range.getCell(row, 1).getValue();
    if (target === false) {
      continue;
    }

    let name = range.getCell(row, 2).getDisplayValue();
    if (name === "") {
      continue;
    }
    let power = range.getCell(row, 3).getValue();
    let count = range.getCell(row, 4).getValue();
    results.push(custom_format_(config, name, power, count));    
  }

  let html = HtmlService.createTemplateFromFile('index');
  html.data = results;
  let evalHtml = html.evaluate();
  evalHtml.setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(evalHtml, '統制判定用のテキスト表示');
}

function show_result() {
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません');
    return;
  }
  console.log(sheet.getSheetName());

  let config_range = get_config_range_();
  let config = read_config_(config_range);

  let results = [];
  // "A2:D17"
  let range = get_data_range_(sheet);
  let prevNote = "";
  for (let row = 1; row <= MAX_SUCCESS_ELEMENT; row++) {
    let name = range.getCell(row, 2).getDisplayValue();
    if (name === "") {
      continue;
    }
    let power = range.getCell(row, 3).getValue();
    let count = range.getCell(row, 4).getValue();
    let note = range.getCell(row, 2).getNote();
    if (note === "") {
      // 成長なしのテキスト
      results.push(custom_format_(config, name, power, count));
      continue;
    }
    if (prevNote !== note) {
      // 成長、もしくは分割の一つ目のテキスト
      results.push(note);// 成長前、分割前
      prevNote = note;
      results.push("→" + custom_format_(config, name, power, count));// 成長後、分割後1
    } else {
      // 分割の二つ目のテキスト
      results.push("→" + custom_format_(config, name, power, count));// 分割後2
    }
  }

  let html = HtmlService.createTemplateFromFile('index');
  html.data = results;
  let evalHtml = html.evaluate();
  evalHtml.setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(evalHtml, '成長申請用のテキスト表示');
}

// below functions are utilities.
function inputBoxCustum_(guideMessage) {
  let inputValue = Browser.inputBox(guideMessage);
  if (inputValue === 'cancel') {
    inputValue = "";
  }
  return inputValue;
}

function get_data_range_(sheet) {
  // "A4:D19"
  let a1notation = "A" + BASE_ROW +":D" + LAST_AVAILABLE_ROW;
  return sheet.getRange(a1notation);
}

function get_config_range_() {
  let config_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  return config_sheet.getRange("A1:I14");
}

function read_config_(config_range) {
  return {
    pre: config_range.getCell(4, 1).getValue(),
    pre_power: config_range.getCell(4, 3).getValue(),
    post_power: config_range.getCell(4, 5).getValue(),
    pre_count: config_range.getCell(4, 6).getValue(),
    post_count: config_range.getCell(4, 8).getValue(),
    post: config_range.getCell(4, 9).getValue(),
    symbol_as_count: config_range.getCell(6, 1).getValue(),
  }
}

function custom_format_(config, name, power, count) {
  let count_str;
  if (config.symbol_as_count === '') {
    count_str = count === 0 ? "" : `${config.pre_count}${number_half_wide_map[count]}${config.post_count}`;
  } else {
    count_str = count === 0 ? "" : `${config.pre_count}${config.symbol_as_count.repeat(count)}${config.post_count}`;
  }
  return `${config.pre}${name}${config.pre_power}${number_half_wide_map[power]}${config.post_power}${count_str}${config.post}`;
}

function power_up_(max_power, power) {
  return Math.min(max_power, power + 1);
}
