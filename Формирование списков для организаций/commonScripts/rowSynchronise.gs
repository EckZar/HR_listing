//==========================================================================
//========Вспомогательная функция сравнения двух массивов===================
//==========================================================================
function rowsSyncronisation(){ // Синхронизация строк из таблиц регионов в таблицы ИЗИ и Общую ЛК
  for(key in regionsSheets) // Обход массива объектов регионов с ссылками на их главные и ИЗИ таблицы
  {
    let region = regionsSheets[key]; // Берем объект региона
    sync(region.main, region.isi); // Отправляем ID таблиц региона(main) и ИЗИ(isi)
  }
}

function sync(mainId, isiId){ // Скрипт использует данные из таблицы региона для синхронизации строк с строками из таблиц ИЗИ и ЛК

  let main = SpreadsheetApp.openById(mainId); // Заходим в главную таблицу

  main.getSheets() // берем список листов в виде массива объектов
  .map( arr => arr.getSheetName()) // приводим массив объектов листов в массив с наименованиями листом
  .filter(function(e)
    {
      return e.indexOf("Оформление_V.2")>=0; // Оставляем только листы, у которых в названии есть Оформление_V.2. Это делается только из-за таблицы "Потребность" в которой два листа Сибирь и Урал
    })
  .forEach(function(item) // Для таких листов делаем  обход
    {
      let data = main.getSheetByName(item); // Заходим в лист
      let mainRange = data.getRange(2, 1, data.getLastRow()-1, 15).getValues(); // Берем все данные с листа      
      syncWithIsi(mainRange, isiId, item);                                            //                         и отправляем синхронищироваться с данными в таблице ИЗИ      
      syncWithLk(mainRange);                                                           //                         и отправляем синхронищироваться с данными в общей таблицеЛК
    })

}

function syncWithIsi(isiRangeFromMain, isiId, sheetName){ // Функция синхронизации массива данных строк ИЗИ из главной таблицы, в таблице ИЗИ
  
  let isiSheet = SpreadsheetApp.openById(isiId).getSheetByName(sheetName); // Заходим в отдельну таблицу ИЗИ региона
  let rangeInIsiSheet = isiSheet.getRange(2, 1, isiSheet.getLastRow()-1, 15).getValues(); // Берем все данные что там есть

  for(let i = 0; i<isiRangeFromMain.length; i++) // начинаем обход по массиву данных ИЗИ из главной таблицы
  {
    for(let j = 0; j<rangeInIsiSheet.length; j++) // начинаем обход по массиву данных ИЗИ из отдельной таблицы для ИЗИ
    {
      if(isiRangeFromMain[i][14] === rangeInIsiSheet[j][14]) // Отдельно сравниваем UUID из каждой строки. If TRUE Идем ниже и делавем полную проверку строк
      { 
        let arrOne = isiRangeFromMain[i].slice(0,-3);
        let arrTwo = rangeInIsiSheet[j].slice(0,-3);
        let diff = compareArrsV2(arrOne, arrTwo);
        if(diff) // проверка расхождений в ячейках строки
        { 
          Logger.log("ISI - " + " " + isiRangeFromMain[i][14] + " <> " + rangeInIsiSheet[j][14] + enterString +
          "fromMain - " + isiRangeFromMain[i] + enterString +
          "fromIsi - " + rangeInIsiSheet[j] + enterString +
          "wrongCell - " + diff + enterString +
          "==========================================================================================================");
          isiSheet.getRange(j+2, diff).setValue(arrOne[diff-1]); // Если есть отличия, заменяем на новую строку из общей таблицы
          // Блок перемещения строки в поле соответствующих дат
          if(diff == 2)
          {
            let pos = getPosV2(arrOne[diff-1], isiSheet) != (j+2) ? getPosV2(arrOne[diff-1], isiSheet) : getPosV3(arrOne[diff-1], isiSheet);
            Logger.log(pos + " <> " + (j+2));
            isiSheet.moveRows(isiSheet.getRange((j+2) + ":" + (j+2)), pos);
          }
        }
        break;
      }
    }
  } 

}

function syncWithLk(lkRangeFromMain){ // Функция синхронизации массива данных строк ЛК из главной таблицы, в общей таблице ЛК

  let rangeInLkSheet = lkSheet.getRange(2, 1, lkSheet.getLastRow()-1, 15).getValues(); // Заходим в общую таблицу ЛК региона и берем все данные что там есть

  for(let i = 0; i<lkRangeFromMain.length; i++) // начинаем обход по массиву данных ЛК из главной таблицы
  {
    for(let j = 0; j<rangeInLkSheet.length; j++) // начинаем обход по массиву данных ЛК из общей таблицы ЛК
    {
      if(lkRangeFromMain[i][14] === rangeInLkSheet[j][14]) // Отдельно сравниваем UUID из каждой строки. If TRUE Идем ниже и делавем полную проверку строк
      {
        let arrOne = lkRangeFromMain[i].slice(0,-3);
        let arrTwo = rangeInLkSheet[j].slice(0,-3);
        let diff = compareArrsV2(arrOne, arrTwo);
        if(diff) // проверка расхождений в ячейках строки
        { 
          Logger.log("LK - "  + " " + lkRangeFromMain[i][14] + " <> " + rangeInLkSheet[j][14] + enterString +
          "fromMain - " + lkRangeFromMain[i] + enterString +
          "fromLK - " + rangeInLkSheet[j] + enterString +
          "wrongCell - " + diff + enterString +
          "==========================================================================================================");
          lkSheet.getRange(j+2, diff).setValue(arrOne[diff-1]); // Если есть отличия, заменяем на новую строку из общей таблицы
          // Блок перемещения строки в поле соответствующих дат
          if(diff == 2)
          {
            let pos = getPosV2(arrOne[diff-1], lkSheet) != (j+2) ? getPosV2(arrOne[diff-1], lkSheet) : getPosV3(arrOne[diff-1], lkSheet);
            Logger.log(pos + " <> " + (j+2));
            lkSheet.moveRows(lkSheet.getRange((j+2) + ":" + (j+2)), pos);
          }
        }
        break;
      }
    }
  } 

}

//==========================================================================
//========Вспомогательная функция сравнения двух массивов===================
//==========================================================================
function compareArrsV2(arrOne, arrTwo){

  for(let i = 0; i<arrOne.length; i++)
  { 
    if(i == 12 || i == "12"){ break;} // Сравнение по статусу не будет проверятся и обрываем проверку на этому месте что бы не проверять комментарии
    
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
      return i+1; // Возвращаем номер столбца который надо заменить
    }
    
  }
  return false;
}

//=============================================================================
//========Вспомогательные функции поиска позиции для сортировки по дате========
//=============================================================================
function getPosV2(date, sheet){

  date = new Date(date).getTime();

  let dates = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues();

  for(let i = 0; i<dates.length; i++)
  {
    let a = new Date(dates[i]).getTime();
    if(date <= a)
    {
      return i+2;
    }
  }
  Logger.log(sheet.getLastRow()+1);
  return sheet.getLastRow()+1;
}

function getPosV3(date, sheet){

  date = new Date(date).getTime();

  let dates = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues();

  for(let i = dates.length; i>0; i--)
  {
    let a = new Date(dates[i]).getTime();
    if(date >= a)
    {
      return i+3;
    }
  }
  Logger.log(sheet.getLastRow()+1);
  return sheet.getLastRow()+1;
}
