//============================================================================================
//======Функция сбора новых строк из таблиц регионов и сбор в общую таблицу ЛК================
//============================================================================================
function bulkUpdateSheet(){

  Logger.log("Загрузка новых строк");

  let uuids = lkSheet.getRange(2, 16, lkSheet.getLastRow()-1, 1).getValues()
            .map(function(arr)
                  { 
                    return arr[0]
                  });

 for(key in regionsSheets) // Обход массива объектов регионов с ссылками на их главные и ИЗИ таблицы
  {
    let region = regionsSheets[key]; // Берем объект региона
    let sheet = SpreadsheetApp.openById(region.main); // Открываем таблицу по main ID
    let sheets = sheet.getSheets(); // Берем список листов
    
    sheets.forEach(function(jtem, j) // Начинаем обход по всем листам в поисках если лист содержит "Оформление_V.2"
    {
      let sheetName = jtem.getName();

      if(sheetName.indexOf("Оформление_V.2")>=0) // Если название листа содержит "Оформление_V.2"
      {
        let main = sheet.getSheetByName(sheetName); // Заходим в лист "Оформление_V.2"

        let values = main.getRange(2, 1, main.getLastRow()-1, 16).getValues() // Берем все данные в таблице
                          .filter(function(e)
                                  {
                                    return e[12].indexOf("ЛК")>=0 && e[15] != "" && uuids.indexOf(e[15])<0; // Оставляем только строки ЛК с UUID и сотавляем строки с теми UUID, которые не нашлись в общей ЛК
                                  })
                          .map(function(arr) // Добавляем в конец массивов ID таблицы
                                  {
                                    arr.push(region.main);
                                    arr.push(main.getSheetId());
                                    return arr; 
                                  });
        try
        {                         
          lkSheet.getRange(lkSheet.getLastRow()+1, 1, values.length, 18).setValues(values);          
          Logger.log("Вставлено " + values.length + " строк из таблицы " + sheet.getName()); 
        }
        catch(e){Logger.log("Таблица: "); Logger.log(e)}
      }
    });
  }
}


//=============================================================================
//========Вспомогательные функции поиска позиции для сортировки по дате========
//=============================================================================
function getPos(date){

  date = new Date(date).getTime();

  let dates = lkSheet.getRange(2, 2, lkSheet.getLastRow()-1, 1).getValues();

  for(let i = 0; i<dates.length; i++)
  {
    let a = new Date(dates[i]).getTime();
    if(date <= a)
    {
      lkSheet.insertRows(i+2, 1);
      return i+2;
    }
  }
  Logger.log(lkSheet.getLastRow()+1);
  return lkSheet.getLastRow()+1;
}






