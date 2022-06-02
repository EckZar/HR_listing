//============================================================================================
//====фуникция проверки статуса оформлен и отправки такого статуса в основную таблицу=========
//============================================================================================
function check(){

  let active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  
  let cell = active.getA1Notation();  
  let value = active.getValue();  
  
  if(cell.indexOf("N">=0) && value === "Оформлен")
  {
    let range = cell.replace(/N/, "I");
    let uuid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление").getRange(range).getValue();
    
    let main = SpreadsheetApp.openById("1ThOToLXcUNPqVzbhXWxg269u63P3iJzgXZJ3J_Wb_Vg").getSheetByName("Оформление");
    
    let mainUuuidsRange = main.getRange(2, 9, main.getLastRow()-1, 1).getValues();
    
    for(let i = 0; i < mainUuuidsRange.length; i++)
    {
      if(mainUuuidsRange[i][0] === uuid)
      {
        main.getRange(i+2,14).setValue("Оформлен");
        return;
      }
    }
  }
}

//============================================================================================
//================фуникция проверки строк по UUID и удаление не совпадений====================
//============================================================================================
function checkByUUID(){

  let mainSheet = SpreadsheetApp.openById("1ThOToLXcUNPqVzbhXWxg269u63P3iJzgXZJ3J_Wb_Vg").getSheetByName("Оформление");
  let mainUUIDs = mainSheet.getRange(2, 9, mainSheet.getLastRow()-1, 1).getValues();

  mainUUIDs = mainUUIDs.filter(function(e){ return e[0] !=""}).map(function(arr){ return arr[0] });

  let isiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление");
  let isiUUIDs = isiSheet.getRange(2, 9, isiSheet.getLastRow()-1, 1).getValues();

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








