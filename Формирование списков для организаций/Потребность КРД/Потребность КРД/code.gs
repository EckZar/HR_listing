//===================================================
//====Скрипт создает UUID для каждой строки ИЗИ======
//===================================================
function fillUUID(){ 

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление");

  let params = main.getRange(2, 2, main.getLastRow()-1, 15).getValues(); // Берем список имен для проверки по наличию

  params = params.map(function(arr, i)
                      { 
                        return [arr[1], // Поле "ФИО" - have to return true
                                arr[7], // Поле "UUID" - have to return false
                                arr[13], // Поле "ЛК/ИЗИ" - have to contain "ИЗИ"
                                arr[3], // Поле "ГОРОД" - have to return true
                                "I"+(i+2)];
                      })
                  .filter(function(e)
                          { 
                            return e[0] != "" && e[1] === "" && e[2].indexOf("ИЗИ")>=0 && e[3] != ""; 
                          })
                  .forEach(function(item)
                          {
                            main.getRange(item[4]).setValue(Utilities.getUuid());
                          })
 
}

//===================================================
//====Скрипт создает UUID для каждой строки ИЗИ======
//===================================================
function fillUUIDV2(){ 

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление_V.2");

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
function transfer(){ 
  
  let newSheet = SpreadsheetApp.openById(isiSheetId).getSheetByName("Оформление");
  let uuids = newSheet.getRange(2, 9, newSheet.getLastRow()-1, 1).getValues().map(function(arr){return arr[0]});

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление");
  let range = main.getRange(2, 1, main.getLastRow()-1, 16).getValues();

  range = range.filter(function(e){ return e[8] !=""; })
               .filter(function(e){ return uuids.indexOf(e[8])<0; })
               .forEach(function(item){ newSheet.getRange(getPos(item[0], isiSheetId), 1, 1, 16).setValues([item]); })
}

//===================================================
//=====Скрипт переноса новых строк в таблицу ИЗИ=====
//===================================================
function transferV2(){
  
  let newSheet = SpreadsheetApp.openById("1HgF4z0jmfjRLYgw-htZ6rbL0C2M_V2Zr3BEa1MyxcKI").getSheetByName("Оформление_V.2");
  let uuids = newSheet.getRange(2, 15, newSheet.getLastRow()-1, 1).getValues().map(function(arr){return arr[0]});

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление_V.2");
  let range = main.getRange(2, 1, main.getLastRow()-1, 15).getValues();

  range = range.filter(function(e){ return e[14] != "" && e[11].indexOf("ИЗИ")>=0; })
               .filter(function(e){ return uuids.indexOf(e[14])<0; })
               .forEach(function(item){newSheet.getRange(getPosV2(item[1], "1HgF4z0jmfjRLYgw-htZ6rbL0C2M_V2Zr3BEa1MyxcKI"), 1, 1, 15).setValues([item]); })
}

//===================================================
//=========Скрипт удаления дубликаотов UUID==========
//===================================================
function removeUUIDDuplicates(){
  
  let isiSheet = SpreadsheetApp.openById(isiSheetId).getSheetByName("Оформление");
  let isiRange = isiSheet.getRange(2, 1, isiSheet.getLastRow()-1, 9).getValues();

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление");
  let range = main.getRange(2, 9, main.getLastRow()-1, 1).getValues().filter(function(e){ return e[0] != ""}).map(function(arr){ return arr[0]});
 
  let positions = main.getRange(2, 1, main.getLastRow()-1, 9).getValues() // Возвращаем массив объектов с параметрами uuid, датой и строкой в таблице
                    .map(function(arr, i)
                    {                    
                      return {
                        "uuid":arr[8],
                        "date": new Date(arr[0]),
                        "count":i+2
                      }
                    })
                    .filter(function(e)
                    { 
                      return e.uuid != "";
                    }); 

  let duplicates = range.filter(onlyUnique) // Возвращаем массив объектов с параметрами uuid и строкой в таблице для uuid которые повторяются в positions более одного раза
                    .map(function(arr){                    
                      return {
                        "uuid":arr,
                        "count":range.filter(function(e){return e === arr}).length
                      }
                    })
                    .filter(function(e){
                      return e.count>1;
                    });
  


  duplicates.forEach(function(item){ // uuid которые встретились более одного раза. Идет проверка по датам из первого столбца главной и ИЗИ таблиц. Если даты различны uuid удаляется

    let date = isiRange.filter(function(e){ return e[8] === item.uuid});
    date = new Date(date[0][0]);   
    dateOne = date.getDate() + "-" + date.getMonth() + "-" + date.getFullYear();
    positions.filter(function(e)
                      { 
                        let dateTwo = new Date(e.date);
                        dateTwo = dateTwo.getDate() + "-" + dateTwo.getMonth() + "-" + dateTwo.getFullYear();

                        return e.uuid === item.uuid && dateOne != dateTwo;
                      })
              .forEach(function(jtem)
                      {
                        Logger.log(jtem);
                        main.getRange(jtem.count, 9).setValue("");
                      });
   
  });
}

//===================================================
//=========Скрипт удаления дубликатов UUID===========
//===================================================
function removeUUIDDuplicatesV2(){

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление_V.2");
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
                        main.getRange(jtem[1], 15).setValue("");
                      })
                    });  

  
}

//====================================================================
//====Скрипт проверки статуса "оформлен" из ИЗИ в основной таблице====
//====================================================================
function check(){

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление");
  let isiSheet = SpreadsheetApp.openById(isiSheetId).getSheetByName("Оформление");

  let range = main.getRange(2, 9, main.getLastRow()-1, 8).getValues();
  let isiRange = isiSheet.getRange(2, 9, main.getLastRow()-1, 8).getValues();

  isiRange = isiRange.filter(function(e){ return e[7] === "Оформлен" });

  for(i in isiRange)
  {
    for(let j = 0 ; j<range.length ; j++)
    {
      if(isiRange[i][0] === range[j][0] && range[j][7] != "Оформлен")
      {
        main.getRange(j+2, 16).setValue("Оформлен");
        break;
      }
    }
  }
}

//=========================================================================
//====Скрипт проверки статуса из ИЗИ и ЛК в основной таблице===============
//=========================================================================
function checksV2(){
  checkV2("1HgF4z0jmfjRLYgw-htZ6rbL0C2M_V2Zr3BEa1MyxcKI"); // ИЗИ
  checkV2("18ivjv9-ueqfF6CZl82mChEh0_JCHNHe79bwswBh7fZA"); // ЛК
}
function checkV2(sheetId){

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление_V.2");  
  let sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Оформление_V.2");

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
function sync(){

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление");
  let isiSheet = SpreadsheetApp.openById(isiSheetId).getSheetByName("Оформление");

  let range = main.getRange(2, 1, main.getLastRow()-1, 16).getValues().filter(function(e){ return e[8] != ""});
  let isiRange = isiSheet.getRange(2, 1, main.getLastRow()-1, 16).getValues();


  for(let i = 0; i<range.length; i++)
  {
    for(let j = 0; j<isiRange.length; j++)
    {
      if(range[i][8] === isiRange[j][8])
      {
        if(!compareArrs(range[i], isiRange[j])) // проверка расхождений в ячейках строки
        {     
          isiSheet.getRange(j+2, 1, 1, 16).setValues([range[i]]); // Если есть отличия, заменяем на новую строку из общей таблицы
        }
        break;
      }
    }
  }
}

//==========================================================================
//========Скрипт синхронизации ИЗИ строк для ИЗИ таблицы====================
//==========================================================================
function syncsV2(){

  syncV2("1HgF4z0jmfjRLYgw-htZ6rbL0C2M_V2Zr3BEa1MyxcKI"); // ИЗИ
  syncV2("18ivjv9-ueqfF6CZl82mChEh0_JCHNHe79bwswBh7fZA"); // ЛК

}
function syncV2(sheetId){

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление_V.2");
  let sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Оформление_V.2");

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
function compareArrs(arrOne, arrTwo){

  for(let i = 1; i<arrOne.length; i++)
  { 
    if(i == 15 || i == "15"){ continue;} // Сравнение по статусу не будет проверятся.
    if(typeof(arrOne[i]) != "object")
    {
      if(arrOne[i] != arrTwo[i])
      { 
        return false;
      }
    }
  }
  return true;
}

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
function getPos(date, sheetId){

  date = new Date(date).getTime();


  let isiSheet = SpreadsheetApp.openById(sheetId).getSheetByName("Оформление");
  let dates = isiSheet.getRange(2, 1, isiSheet.getLastRow()-1, 1).getValues();

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

function getPosV2(date, sheetId){

  date = new Date(date).getTime();

  let isiSheet = SpreadsheetApp.openById(sheetId).getSheetByName("Оформление_V.2");
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





