//===================================================
//====Скрипт создает UUID для каждой строки ИЗИ======
//===================================================
function fillUUIDV2s(){
  fillUUIDV2("Урал Оформление_V.2");
  fillUUIDV2("Сибирь Оформление_V.2");
}
function fillUUIDV2(sheetName){ 

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  let params = main.getRange(2, 1, main.getLastRow()-1, 15).getValues(); // Берем список имен для проверки по наличию

  params = params.map(function(arr, i)
                      { 
                        return [arr[2], // Поле Рекрутер - have to return true
                                arr[3], // Поле "ФИО кандидата" - have to return true
                                arr[11], // Поле ЛК/ИЗИ - have to return true
                                arr[14], // Поле UUID - have to return false
                                "O"+(i+2)];
                      })
                  .filter(function(e)
                          { 
                            return e[0] != "" && e[1] !="" && e[2] != "" && e[3] === ""; 
                          })
                  .forEach(function(item)
                          {
                            Logger.log(item);
                            main.getRange(item[4]).setValue(Utilities.getUuid());
                          })
 
}

//===================================================
//=====Скрипт переноса новых строк в таблицу ИЗИ=====
//===================================================
function transferV2s(){
  transferV2("Урал Оформление_V.2");
  transferV2("Сибирь Оформление_V.2");
}
function transferV2(sheetName){
  
  let newSheet = SpreadsheetApp.openById(isiSheetId).getSheetByName(sheetName);
  let uuids = newSheet.getRange(2, 15, newSheet.getLastRow()-1, 1).getValues().map(function(arr){return arr[0]});

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let range = main.getRange(2, 1, main.getLastRow()-1, 15).getValues();

  range = range.filter(function(e){ return e[14] != "" && e[11].indexOf("ИЗИ")>=0; })
               .filter(function(e){ return uuids.indexOf(e[14])<0; })
               .forEach(function(item){newSheet.getRange(getPosV2(item[1], isiSheetId, sheetName), 1, 1, 15).setValues([item]); })
}

//===================================================
//=========Скрипт удаления дубликатов UUID===========
//===================================================
function removeUUIDDuplicatesV2s(){
  removeUUIDDuplicatesV2("Урал Оформление_V.2");
  removeUUIDDuplicatesV2("Сибирь Оформление_V.2");
}
function removeUUIDDuplicatesV2(sheetName){

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let range = main.getRange(2, 15, main.getLastRow()-1, 1).getValues().filter(function(e){ return e[0] != ""});
  let posses = main.getRange(2, 15, main.getLastRow()-1, 1).getValues().map(function(arr, i){ return [arr[0], i+2]}).filter(function(e){ return e[0] != ""});
  let inLine = range.map(function(arr){ return arr[0]});

  let duplicates = inLine.filter(onlyUnique) // Возвращаем массив объектов с параметрами uuid и строкой в таблице для uuid которые повторяются в positions более одного раза
                    .map(function(arr){                    
                      return {
                        "uuid": arr,
                        "count": inLine.filter(function(e){return e === arr}).length
                      }
                    })
                    .filter(function(e){
                      return e.count>1;
                    })
                    .forEach(function(item, i)
                    {                      
                      posses.filter(function(e)
                      {
                        return e[0] == item.uuid
                      })
                      .forEach(function(jtem, j)
                      {
                        Logger.log(jtem);
                        main.getRange(jtem[1], 15).setValue("");
                      })
                    });  

  
}

//=========================================================================
//====Скрипт проверки статуса из ИЗИ и ЛК в основной таблице===============
//=========================================================================
function checksV2(){
  checkV2(isiSheetId, "Урал Оформление_V.2", "Урал Оформление_V.2"); // ИЗИ "Урал Оформление_V.2"
  checkV2(isiSheetId, "Сибирь Оформление_V.2", "Сибирь Оформление_V.2"); // ИЗИ "Сибирь Оформление_V.2"

  checkV2(commonLkSheetId, "Урал Оформление_V.2", "Оформление_V.2"); // ЛК
  checkV2(commonLkSheetId, "Сибирь Оформление_V.2", "Оформление_V.2"); // ЛК
}
function checkV2(sheetId, sheetNameHere, sheetNameThere){

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameHere);  
  let sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetNameThere);

  let mRange = main.getRange(2, 12, main.getLastRow()-1, 4).getValues();
  let sRange = sheet.getRange(2, 12, sheet.getLastRow()-1, 4).getValues();  

  for(let i = 0 ; i<sRange.length ; i++)
  {
    for(let j = 0 ; j<mRange.length ; j++)
    {
      if(sRange[i][3] === mRange[j][3] && mRange[j][1] != sRange[i][1])
      {
        Logger.log(sRange[i][3] + " <> " + mRange[j][3]);
        Logger.log(sRange[i][1] + " <> " + mRange[j][1]);
        Logger.log("=================================================================");
        main.getRange(j+2, 13).setValue(sRange[i][1]);
        break;
      }
    }
  }
}

//==========================================================================
//========Скрипт синхронизации ИЗИ строк для ИЗИ таблицы====================
//==========================================================================
function syncV2s(){

  syncV2(isiSheetId, "Урал Оформление_V.2", "Урал Оформление_V.2"); // ИЗИ "Урал Оформление_V.2"
  syncV2(isiSheetId, "Сибирь Оформление_V.2", "Сибирь Оформление_V.2"); // ИЗИ "Сибирь Оформление_V.2"

  syncV2(commonLkSheetId, "Урал Оформление_V.2", "Оформление_V.2"); // ЛК
  syncV2(commonLkSheetId, "Сибирь Оформление_V.2", "Оформление_V.2"); // ЛК

}
function syncV2(sheetId, sheetNameHere, sheetNameThere){

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameHere);
  let sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetNameThere);

  let range = main.getRange(2, 1, main.getLastRow()-1, 15).getValues().filter(function(e){ return e[14] != ""});
  
  let sRange = sheet.getRange(2, 1, sheet.getLastRow()-1, 15).getValues(); 

  for(let i = 0; i<range.length; i++)
  {
    for(let j = 0; j<sRange.length; j++)
    {
      if(range[i][14] === sRange[j][14])
      {
          
        if(!compareArrsV2(range[i], sRange[j])) // проверка расхождений в ячейках строки
        {               
          Logger.log(range[i]);
          Logger.log(sRange[j]);
          Logger.log("==========================================================================================================");
          sheet.getRange(j+2, 1, 1, 15).setValues([range[i]]); // Если есть отличия, заменяем на новую строку из общей таблицы
        }
        break;
      }
    }
  } 
 
  sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).sort(2);

}

//==========================================================================
//========Вспомогательная функция сравнения двух массивов===================
//==========================================================================
function compareArrsV2(arrOne, arrTwo){

  for(let i = 0; i<arrOne.length; i++)
  { 
    if(i == 12 || i == "12"){ continue;} // Сравнение по статусу не будет проверятся.
    
    try{
      var a = arrOne[i].toString().replace(/\s/g,"").toLowerCase();
      var b = arrTwo[i].toString().replace(/\s/g,"").toLowerCase();
    }
    catch(e)
    { 
      var a = arrOne[i];
      var b = arrTwo[i];

      Logger.log(e);
      Logger.log(a);
      Logger.log(b);
      Logger.log("====================================================");
    }
    if(a != b)
    { 
      return false;
    }
    
  }
  return true;
}

//=============================================================================
//========Вспомогательные функции поиска позиции для сортировки по дате========
//=============================================================================

function getPosV2(date, sheetId, sheetName){

  date = new Date(date).getTime();

  let isiSheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  let dates = isiSheet.getRange(2, 2, isiSheet.getLastRow()-1, 1).getValues();

  for(let i = 0; i<dates.length; i++)
  {
    let a = new Date(dates[i]).getTime();
    if(date <= a)
    {
      isiSheet.insertRows(i+2, 1);
      return i+2;
    }
  }
  
  return isiSheet.getLastRow()+1;
}


//=============================================================================
//========Вспомогательные функции поиска дубликатов в массиве==================
//=============================================================================
function onlyUnique(value, index, self) {
    return self.indexOf(value) === index;
}





