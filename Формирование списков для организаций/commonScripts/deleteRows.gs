//=============================================================================
//========Функция удаления сотрудников ИЗИ из общей таблицы ЛК=================
//=============================================================================
function deleteIsiRows(){  
 
  let range = lkSheet.getRange(2, 13, lkSheet.getLastRow()-1, 1).getValues();

  for(let i = range.length-1; i >= 0 ; i--)
  { 
    try
    {
      if(range[i][0].indexOf("ИЗИ")>=0)
      {
        Logger.log(lkSheet.getRange(i+2, 1, 1, 16).getValues());
        lkSheet.deleteRow(i+2);
      }
    }
    catch(e){Logger.log(e); Logger.log(range[i])};
  }

}

//=================================================================================================================
//========Функция поиска строк по UUID из ЛК в главных таблицах и удаление если ничего не нашлось==================
//=================================================================================================================
function removeUnrecognizedUUID(){

  let forDelete = [];

  let range = lkSheet.getRange(2, 16, lkSheet.getLastRow()-1, 2).getValues().map(function(arr, i){ return [arr[0], arr[1], i+2]});

  sheetsIds.forEach(function(item, i)
  {

    let arr = range.filter(function(e)
                  {
                    return e[1] === item
                  });

    let sheet = SpreadsheetApp.openById(item);
    let sheets = sheet.getSheets();

    let vals = [];

    sheets.forEach(function(name, n)
                    {
                      let sheetName = name.getName();
                      if(sheetName.indexOf("Оформление_V.2")>=0)
                      {
                        let data = sheet.getSheetByName(sheetName);
                        let temp = data.getRange(2, 16, data.getLastRow()-1, 1).getValues().map(arr => arr[0]);
                        vals = vals.concat(temp);
                      }
                    })

    arr.forEach(function(jtem, j)
                        { 
                          !~vals.indexOf(jtem[0]) ? forDelete.push(jtem[2]) : '';
                        })
  })

  forDelete.sort(function (a, b) 
    {
      return b - a;
    })
    .forEach(function(item, i)
    {
      Logger.log(item);
      lkSheet.deleteRow(item);
    });
}


function deleteEmptyRow(){

  let range = lkSheet.getRange(1, 16, lkSheet.getLastRow(), 1).getValues();
  for(let i = range.length-1; i >= 0 ; i--)
  { 
    if(range[i][0] === "")
    {
      Logger.log(lkSheet.getRange(i+1, 1, 1, 16).getValues());
      Logger.log(i+1);
      lkSheet.deleteRow(i+1);
    }
  }
}

function anotherEmpty(){ // Ищем строки где стоит Оформлен или еще что, но остальная строка пустая.

  for(key in regionsSheets)
  {
    let mainId = regionsSheets[key]["isi"];    
    let main = SpreadsheetApp.openById(mainId).getSheetByName("Оформление_V.2");
    let range = main.getRange(2, 14, main.getLastRow()-1, 3).getValues().map((obj, i) => [obj[0], obj[2], "O" + (i+2)]).filter(e => e[1] == "" && e[0] !="");
    Logger.log(range);
    Logger.log(mainId);
    // let isi = regionsSheets[key][isi];
  }
}


//==================================================================================================
//================================фуникция удаления дубликатов======================================
//==================================================================================================

function removeDuplicatesByUUID(){  
  Logger.log("Удаление дубликатов")
  lkSheet.getRange(2, 1, lkSheet.getLastRow()-1, 18).removeDuplicates([16]);
}


















