class ConvertString{
  /**
   * @param {Array.string} Value of the sheet.
   * @param {number} Index of the column.
   */
  constructor(targetValues, targetColIdx){
    this.targetValues = targetValues;
    this.targetColIdx = targetColIdx;
  }
  /**
  * Remove all spaces.
  */
  removeAllSpace(targetStr){
    return targetStr.replace(/\s+/g, '');
  }
  /**
   * Removes whitespace from all strings in the specified column.
   * @return {Array.string}
   */
  removeAllSpaceTargetCol(){
    let temp = this.targetValues;
    temp.forEach(x => x[this.targetColIdx] = this.removeAllSpace(x[this.targetColIdx]));
    return temp;
  }
  /**
   * Converts the string in the column specified by targetColIdx to lowercase.
   * @return {Array.string}
   */
  convertToLower(){
    let temp = this.targetValues;
    temp.forEach(x => x[this.targetColIdx] = x[this.targetColIdx].toLowerCase());
    return temp;
  }
}
class CheckSinetAndShisan{
  constructor(){
    this.ssSinet = SpreadsheetApp.openByUrl(PropertiesService.getScriptProperties().getProperty('sinetUrl'));
    this.ssShisankanri = SpreadsheetApp.openByUrl(PropertiesService.getScriptProperties().getProperty('shisankanriUrl'));
    this.ssOutput = SpreadsheetApp.openByUrl(PropertiesService.getScriptProperties().getProperty('outputSheetUrl'));
    this.sinetTargetSheetIdx = 0;
    this.sinetTerminalIdx = 4;
    this.sinetHaikiIdx = 6;
    this.sinetUsernameIdx = 2;
    this.shisanTerminalIdx = 0;
    this.shisanCategoryIdx = 1;
    this.shisanStatusIdx = 4;
    this.shisanUsernameIdx = 23;
  }
  init(){
    const targetSheets = this.ssOutput.getSheets();
    const sheetcount = 4 - targetSheets.length;
    if (sheetcount > 0){
      for (let i = 1; i <= sheetcount; i++){
        this.ssOutput.insertSheet();
      }
    }
    const targetSheets2 = this.ssOutput.getSheets();
    const sheetname = ['SINET接続許可に存在しない端末', 'SINET接続申請内で重複している端末', '使用者名が異なる端末', '機器一覧に存在しない端末'];
    for (let i = 0; i <= 3; i++){
      targetSheets2[i].setName(sheetname[i]);
    }
  }
  outputValues(outputSheetIdx, outputValue){
    // outputValueの値を指定したシートに出力する。一番左のシートはoutputSheetIdx == 0。
    const outputSheet = this.ssOutput.getSheets()[outputSheetIdx];
    outputSheet.clear();
    outputSheet.getRange(1, 1, outputValue.length, outputValue[0].length).setValues(outputValue);
  }
  getUsingSINETRegistrationTerminalValues(){
    // SINET接続許可の廃棄日が入っていないレコードを返す。
    return this.ssSinet.getSheets()[this.sinetTargetSheetIdx].getDataRange().getValues().filter((x, idx) => x[this.sinetHaikiIdx] == '' || idx == 0);
  }
  getUsingShisanTerminalValues(){
    // 資産管理の廃棄済みでないデスクトップPC、ノートPCを返す。
    return this.ssShisankanri.getSheetByName('機器一覧').getDataRange().getValues().filter((x, idx) => ((x[this.shisanCategoryIdx] == 'デスクトップPC' || x[this.shisanCategoryIdx] == 'ノートPC') && x[this.shisanStatusIdx] != '破棄済') || idx == 0);
  }
  removeAllSpaceSinetUsername(){
    const ClassConvertString = new ConvertString(this.getUsingSINETRegistrationTerminalValues(), this.sinetUsernameIdx);
    return ClassConvertString.removeAllSpaceTargetCol();
  } 
  removeAllSpaceShisanUsername(){
    const ClassConvertString = new ConvertString(this.getUsingShisanTerminalValues(), this.shisanUsernameIdx);
    return ClassConvertString.removeAllSpaceTargetCol();
  } 
  convertToLowerSinetTerminalName(){
    // SINET接続許可の端末名を小文字に変換する。
    const ClassConvertString = new ConvertString(this.getUsingSINETRegistrationTerminalValues(), this.sinetTerminalIdx);
    return ClassConvertString.convertToLower();
  }
  convertToLowerShisanTerminalName(){
    // 資産管理の端末名を小文字に変換する。
    const ClassConvertString = new ConvertString(this.getUsingShisanTerminalValues(), this.shisanTerminalIdx);
    return ClassConvertString.convertToLower();
  }
  extractSINETUnregisteredTerminal(){
    // 資産管理の機器一覧とSINET接続許可を突き合わせし、SINET接続許可に存在しない端末を抽出する。
    let sinet = this.convertToLowerSinetTerminalName();
    let shisankanri = this.convertToLowerShisanTerminalName()
    const target = shisankanri.map(x => {
      const temp = sinet.filter(y => x[this.shisanTerminalIdx] == y[this.sinetTerminalIdx]);
      const temp2 = temp.length == 0 ? sinet.filter(y => x[this.shisanTerminalIdx] + '.local' == y[this.sinetTerminalIdx]) : temp;
      return temp2.length == 0 ? x : null;
    }).filter(x => x);
    this.outputValues(0, target);
  }
  getDuplicateSINETRegistrationTerminal(){
    // SINET接続申請内で重複しているコンピューター名を抽出する
    const sinet = this.convertToLowerSinetTerminalName(this.getUsingSINETRegistrationTerminalValues());
    let temp = sinet.map(x => x[parseInt(this.sinetTerminalIdx)]);
    const terminalNameList = Array.from(new Set(temp));
    let duplicateTerminal = terminalNameList.map(x => {
      const temp = sinet.filter(y => y[parseInt(this.sinetTerminalIdx)] == x);
      return temp.length > 1 ? temp[0] : null;
    }).filter(x => x);
    if (duplicateTerminal.length == 0){
      duplicateTerminal = [['対象レコードなし']];
    }
    this.outputValues(1, duplicateTerminal);
  }
  getUserDifference(){
    // 資産管理の機器一覧とSINET接続許可を突き合わせし、使用者名が異なる端末を抽出する。
    const sinet = this.removeAllSpaceSinetUsername();
    const shisan = this.removeAllSpaceShisanUsername();
    const targetTerminals = sinet.map(x => {
      let temp = shisan.filter(y => y[this.shisanTerminalIdx] == x[this.sinetTerminalIdx]);
      let temp2 = (temp.length > 0) && (temp[0][this.shisanUsernameIdx] != x[this.sinetUsernameIdx]) ? [x[this.sinetTerminalIdx], x[this.sinetUsernameIdx], temp[0][this.shisanUsernameIdx]] : null;
      return temp2;
    }).filter(x => x);
    this.outputValues(2, targetTerminals);
  }
  extractShisanUnregisteredTerminal(){
    // 資産管理の機器一覧とSINET接続許可を突き合わせし、機器一覧に存在しない端末を抽出する。
    let sinet = this.convertToLowerSinetTerminalName();
    let shisankanri = this.convertToLowerShisanTerminalName()
    const target = sinet.map(x => {
      const temp = shisankanri.filter(y => y[this.shisanTerminalIdx] == x[this.sinetTerminalIdx]);
      const temp2 = temp.length == 0 ? shisankanri.filter(y => y[this.shisanTerminalIdx] + '.local' == x[this.sinetTerminalIdx]) : temp;
      return temp2.length == 0 ? x : null;
    }).filter(x => x);
    this.outputValues(3, target);
  }
}
function getShisanAndSinetData(){
  const ExecCheck = new CheckSinetAndShisan;
  ExecCheck.init();
  ExecCheck.extractSINETUnregisteredTerminal();
  ExecCheck.getDuplicateSINETRegistrationTerminal();
  ExecCheck.getUserDifference();
  ExecCheck.extractShisanUnregisteredTerminal();
}
/**
 * Set the properties.
 * @param none.
 * @return none.
 */
function registerScriptProperty(){
  PropertiesService.getScriptProperties().deleteAllProperties;
  // SINET接続許可依頼書のスプレッドシートのURL
  const sinetUrl = 'https://docs.google.com/spreadsheets/d/...';
  // 資産管理のスプレッドシートのURL
  const shisankanriUrl = 'https://docs.google.com/spreadsheets/d/...';
  // 出力用スプレッドシートのURL
  const outputSheetUrl = 'https://docs.google.com/spreadsheets/d/...';
  PropertiesService.getScriptProperties().setProperty('sinetUrl', sinetUrl);
  PropertiesService.getScriptProperties().setProperty('shisankanriUrl', shisankanriUrl);
  PropertiesService.getScriptProperties().setProperty('outputSheetUrl', outputSheetUrl);
}
