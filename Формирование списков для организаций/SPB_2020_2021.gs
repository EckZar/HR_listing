//===================================================
//====Скрипт создает UUID для каждой строки ИЗИ======
//===================================================
function fillUUID(){ 

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление");

  let params = main.getRange(2, 2, main.getLastRow()-1, 12).getValues(); // Берем список имен для проверки по наличию

  params = params.map(function(arr, i)
                      { 
                        return [arr[0], // Поле ФИО - have to return true
                                arr[7], // Поле UUID - have to return false
                                arr[10], // Поле ЛК/ИЗИ - have to return "ИЗИ"
                                arr[11], // Поле согласование - have to return true
                                "I"+(i+2)];
                      })
                  .filter(function(e)
                          { 
                            return e[0] != "" && e[1] ==="" && e[2] === "ИЗИ" && e[3] != ""; 
                          })
                  .forEach(function(item)
                          {
                            main.getRange(item[4]).setValue(Utilities.getUuid());
                          })
 
}


//===================================================
//=====Скрипт переноса новых строк в таблицу ИЗИ=====
//===================================================
function transfer(){ 
  
  let newSheet = SpreadsheetApp.openById("1QRvwCRKseLpSSh97_OJLqC31m7_ccK4Xxyq-Quj43V4").getSheetByName("Оформление");
  let uuids = newSheet.getRange(2, 9, newSheet.getLastRow()-1, 1).getValues().map(function(arr){return arr[0]});

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление");
  let range = main.getRange(2, 1, main.getLastRow()-1, 15).getValues();

  range = range.filter(function(e){ return e[8] !=""; })
               .filter(function(e){ return uuids.indexOf(e[8])<0; })
               .forEach(function(item){ newSheet.getRange(getPos(item[0]), 1, 1, 15).setValues([item]); })
}

//===================================================
//=========Скрипт удаления дубликаотов UUID==========
//===================================================
function removeUUIDDuplicates(){
  
  let isiSheet = SpreadsheetApp.openById("1QRvwCRKseLpSSh97_OJLqC31m7_ccK4Xxyq-Quj43V4").getSheetByName("Оформление");
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


//====================================================================
//====Скрипт проверки статуса "оформлен" из ИЗИ в основной таблице====
//====================================================================
function check(){

  let main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Оформление");
  let isiSheet = SpreadsheetApp.openById("1QRvwCRKseLpSSh97_OJLqC31m7_ccK4Xxyq-Quj43V4").getSheetByName("Оформление");

  let range = main.getRange(2, 9, main.getLastRow()-1, 6).getValues();
  let isiRange = isiSheet.getRange(2, 9, main.getLastRow()-1, 6).getValues();

  isiRange = isiRange.filter(function(e){ return e[5] === "Оформлен" });

  for(i in isiRange)
  {
    for(let j = 0 ; j<range.length ; j++)
    {
      if(isiRange[i][0] === range[j][0] && range[j][5] != "Оформлен")
      {
        main.getRange(j+2, 14).setValue("Оформлен");
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
  let isiSheet = SpreadsheetApp.openById("1QRvwCRKseLpSSh97_OJLqC31m7_ccK4Xxyq-Quj43V4").getSheetByName("Оформление");

  let range = main.getRange(2, 1, main.getLastRow()-1, 15).getValues().filter(function(e){ return e[8] != ""});
  let isiRange = isiSheet.getRange(2, 1, main.getLastRow()-1, 15).getValues();


  for(let i = 0; i<range.length; i++)
  {
    for(let j = 0; j<isiRange.length; j++)
    {
      if(range[i][8] === isiRange[j][8])
      {
        if(!compareArrs(range[i], isiRange[j])) // проверка расхождений в ячейках строки
        { 
          // let newPos = getPosTwo(range[i][0]);
          // Logger.log(newPos + " <> " + range[i][1]);        
          isiSheet.getRange(j+2, 1, 1, 15).setValues([range[i]]); // Если есть отличия, заменяем на новую строку из общей таблицы
          // isiSheet.moveRows(isiSheet.getRange(j+2,1), newPos);
        }
        break;
      }
    }
  }
}


//==========================================================================
//========Вспомогательная функция сравнения двух массивов===================
//==========================================================================
function compareArrs(arrOne, arrTwo){

  for(let i = 1; i<arrOne.length; i++)
  { 
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


//=============================================================================
//========Вспомогательные функции поиска позиции для сортировки по дате========
//=============================================================================
function getPos(date){

  date = new Date(date).getTime();


  let isiSheet = SpreadsheetApp.openById("1QRvwCRKseLpSSh97_OJLqC31m7_ccK4Xxyq-Quj43V4").getSheetByName("Оформление");
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

function getPosTwo(date){

  date = new Date(date).getTime();


  let isiSheet = SpreadsheetApp.openById("1QRvwCRKseLpSSh97_OJLqC31m7_ccK4Xxyq-Quj43V4").getSheetByName("Оформление");
  let dates = isiSheet.getRange(2, 1, isiSheet.getLastRow()-1, 1).getValues();

  for(let i = 0; i<dates.length; i++)
  {
    let a = new Date(dates[i]).getTime();
    if(date <= a)
    {
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





