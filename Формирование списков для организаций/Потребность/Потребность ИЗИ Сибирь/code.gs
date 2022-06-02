//============================================================================================
//====фуникция проверки статуса оформлен и отправки такого статуса в основную таблицу=========
//============================================================================================
function check(){

  let active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let sheetName = active.getSheetName();
  let val = active.getActiveCell();
  
  let cell = val.getA1Notation();  
  let value = val.getValue();  
  
  if(cell.indexOf("M">=0) && value === "Оформлен")
  {
    let range = cell.replace(/M/, "I");
    let uuid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(range).getValue();
    
    let main = SpreadsheetApp.openById(mainSheetId).getSheetByName(sheetName);
    
    let mainUuuidsRange = main.getRange(2, 9, main.getLastRow()-1, 1).getValues();
    
    for(let i = 0; i < mainUuuidsRange.length; i++)
    {
      if(mainUuuidsRange[i][0] === uuid)
      {
        main.getRange(i+2,13).setValue("Оформлен");
        return;
      }
    }
  }
}

//============================================================================================
//================фуникция проверки строк по UUID и удаление не совпадений====================
//============================================================================================
function checkByUUID(mainSheetId, sheetName){

  let mainSheet = SpreadsheetApp.openById(mainSheetId).getSheetByName(sheetName);
  let mainUUIDs = mainSheet.getRange(2, 9, mainSheet.getLastRow()-1, 1).getValues();

  mainUUIDs = mainUUIDs.filter(function(e){ return e[0] !=""}).map(function(arr){ return arr[0] });

  let isiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let isiUUIDs = isiSheet.getRange(2, 9, isiSheet.getLastRow()-1, 1).getValues();

  isiUUIDs = isiUUIDs.filter(function(e){ return mainUUIDs.indexOf(e[0])<0 });

  Logger.log(isiUUIDs.length);
  Logger.log(isiUUIDs);

}

//============================================================================================
//====фуникция проверки статуса оформлен и отправки такого статуса в основную таблицу=========
//============================================================================================
function checkV2(){

  let active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let sheetName = active.getSheetName();
  let val = active.getActiveCell();
  
  let cell = val.getA1Notation();  
  let value = val.getValue();  

  if(cell.indexOf("M">=0))
  {
    Logger.log(value);
    let range = cell.replace(/M/, "O");
    let uuid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(range).getValue();
    
    let main = SpreadsheetApp.openById(mainSheetId).getSheetByName(sheetName);
    
    let mainUuuidsRange = main.getRange(2, 15, main.getLastRow()-1, 1).getValues();
    
    for(let i = 0; i < mainUuuidsRange.length; i++)
    {
      if(mainUuuidsRange[i][0] === uuid)
      {
        main.getRange(i+2,13).setValue(value);
        return;
      }
    }
  }

  if(cell.indexOf("N">=0))
  {
    Logger.log(value);
    let range = cell.replace(/N/, "O");
    let uuid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(range).getValue();
    
    let main = SpreadsheetApp.openById(mainSheetId).getSheetByName(sheetName);
    
    let mainUuuidsRange = main.getRange(2, 15, main.getLastRow()-1, 1).getValues();
    
    for(let i = 0; i < mainUuuidsRange.length; i++)
    {
      if(mainUuuidsRange[i][0] === uuid)
      {
        main.getRange(i+2,14).setValue(value);
        return;
      }
    }
  }

}

//============================================================================================
//================фуникция проверки строк по UUID и удаление не совпадений====================
//============================================================================================
function checkByUUIDV2s(){
  checkByUUIDV2(mainSheetId, "Сибирь Оформление_V.2");
  checkByUUIDV2(mainSheetId, "Урал Оформление_V.2");
}
function checkByUUIDV2(sheetId, sheetName){

  let mainSheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  let mainUUIDs = mainSheet.getRange(2, 15, mainSheet.getLastRow()-1, 1).getValues();

  mainUUIDs = mainUUIDs.filter(function(e){ return e[0] !=""}).map(function(arr){ return arr[0] });

  let isiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let isiUUIDs = isiSheet.getRange(2, 15, isiSheet.getLastRow()-1, 1).getValues();

  isiUUIDs = isiUUIDs.map(function(arr, i) // Для каждого UUID возвращаем UUID и номер строки
                          { 
                            return [arr[0], i+2]
                          })
                    .filter(function(e) // Для каждого UUID из ИЗИ фильтруем по отсутствию в списке UUIDs из главной таблицы
                          { 
                            return mainUUIDs.indexOf(e[0])<0 
                          });
  
  for(let i = isiUUIDs.length-1; i >= 0 ; i--)
  { 
    Logger.log(isiSheet.getRange(isiUUIDs[i][1], 1, 1, 15).getValues());
    isiSheet.deleteRow(isiUUIDs[i][1]);
  }

}

function deleteLkRows(){
  deleteLkRow("Сибирь Оформление_V.2");
  deleteLkRow("Урал Оформление_V.2");
}
function deleteLkRow(sheetName){
  
  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let range = main.getRange(2, 12, main.getLastRow()-1, 1).getValues();

  for(let i = range.length-1; i >= 0 ; i--)
  { 
    if(range[i][0].indexOf("ЛК")>=0)
    {
      Logger.log(main.getRange(i+2, 1, 1, 15).getValues());
      main.deleteRow(i+2);
    }
  }

}














