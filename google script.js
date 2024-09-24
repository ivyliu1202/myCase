
function doGet(e) {
  //取得參數
  let params = e.parameter; 
  let number = params.number;
  let name = params.name;
  let oneScore = params.oneScore;
  let twoScore = params.twoScore;
  let threeScore = params.threeScore;

 
  //sheet資訊
  let SpreadSheet = SpreadsheetApp.openById("1h0Q84t8Q6GvPLntde07T1s8uL8lOCXFa15Xs8oczj0A");
  let Sheet = SpreadSheet.getSheets()[0];
  let LastRow = Sheet.getLastRow();

  //存入資訊
  Sheet.getRange(LastRow+1, 1).setValue(number);
  Sheet.getRange(LastRow+1, 2).setValue(name);
  Sheet.getRange(LastRow+1, 3).setValue(oneScore);
  Sheet.getRange(LastRow+1, 4).setValue(twoScore);
  Sheet.getRange(LastRow+1, 5).setValue(threeScore);
  if(oneScore>=4 || twoScore>=4 || threeScore>=4){
    Sheet.getRange(LastRow+1,6).setValue("V");
  }
  if(oneScore>=4 && twoScore>=4 && threeScore>=4){
    Sheet.getRange(LastRow+1,7).setValue("V");
    MailApp.sendEmail({
      to: "ivy28644862851586@gmail.com", // 這邊我們直接把取得的 email 帶入
      subject: "【系統信件 請勿回復】痛痛！",
      body: `${name} 你好
      你已經連續痛三天了`
      
    });
  }
  
  //回傳資訊
  return ContentService.createTextOutput("123");
}

function onEdit(e) {
  console.log("987")
  let row = e.range.getRow();
  let col = e.range.getColumn();
  let name = e.source.getActiveSheet().getName();
  let painScoreOne = e.range.getSheet().getRange(1, 1, row+1, 8).getCell(row, 3).getValue();
  let painScoreTwo = e.range.getSheet().getRange(1, 1, row+1, 8).getCell(row, 4).getValue();
  let painScoreThree = e.range.getSheet().getRange(1, 1, row+1, 8).getCell(row, 5).getValue();

  
  if(painScoreOne != ""){
    painScoreOne = parseInt(painScoreOne, 10);
  }
  if(painScoreTwo != ""){
    painScoreTwo = parseInt(painScoreTwo, 10);
  }
  if(painScoreThree != ""){
    painScoreThree = parseInt(painScoreThree, 10);
  }


  if(row > 1 && name == "工作表1"){
    
    
    if(painScoreOne>=4 || painScoreTwo>=4 || painScoreThree>=4){
      SpreadsheetApp.getActiveSheet().getRange(row,6).setValue("V");
    }
    if(painScoreOne>=4 && painScoreTwo>=4 && painScoreThree>=4){
      SpreadsheetApp.getActiveSheet().getRange(row,7).setValue("V");
      let patientName = e.range.getSheet().getRange(1, 1, row+1, 8).getCell(row, 2).getValue();
      MailApp.sendEmail({
        to: "ivy28644862851586@gmail.com", // 這邊我們直接把取得的 email 帶入
        subject: "【系統信件 請勿回復】痛痛！",
        body: `${patientName} 你好
        你已經連續痛三天了`
        
      });
    }
  }
}

function onEditTrigger() {
  const e = {
    range: SpreadsheetApp.getActiveRange()
  };
  console.log("444")
  let row = e.range.getRow();
  let col = e.range.getColumn();
  let name = e.source.getActiveSheet().getName();
  let painScoreOne = e.range.getSheet().getRange(1, 1, row+1, 8).getCell(row, 3).getValue();
  let painScoreTwo = e.range.getSheet().getRange(1, 1, row+1, 8).getCell(row, 4).getValue();
  let painScoreThree = e.range.getSheet().getRange(1, 1, row+1, 8).getCell(row, 5).getValue();

  
  if(painScoreOne != ""){
    painScoreOne = parseInt(painScoreOne, 10);
  }
  if(painScoreTwo != ""){
    painScoreTwo = parseInt(painScoreTwo, 10);
  }
  if(painScoreThree != ""){
    painScoreThree = parseInt(painScoreThree, 10);
  }


  if(row > 1 && name == "工作表1"){
    
    
    if(painScoreOne>=4 && painScoreTwo>=4 && painScoreThree>=4){
      let patientName = e.range.getSheet().getRange(1, 1, row+1, 8).getCell(row, 2).getValue();
      MailApp.sendEmail({
        to: "ivy28644862851586@gmail.com", // 這邊我們直接把取得的 email 帶入
        subject: "【系統信件 請勿回復】痛痛！",
        body: `${patientName} 你好
        你已經連續痛三天了
        `
      });
    }
  }

}





















