//==================================================================================================
//====фуникция проверки статуса оформлен и отправки такого статуса в основную таблицу по ИД=========
//==================================================================================================
function checkStatus(){
  
  let range = mainSheet.getRange(2, 14, mainSheet.getLastRow()-1, 4).getValues(); // Берем список всех значений по столбцам M-P из общей ЛК

  sheetsIds.forEach(function(item) // Начинае обход по всем главным таблицам
  {
    let sheet = SpreadsheetApp.openById(item); // Открываем главную таблицу
    let sheets = sheet.getSheets(); // Берем список листов из главной таблицы
    let temporaryArray = range.filter(function(e) // из общего списка всех значений из Общей ЛК оставляем те, которые по полю ID соответствуют ID текущей таблицы
                              {
                                return e[3] === item
                              });
    
    sheets.forEach(function(name, n) // Начинаем обход по листам в поисках Оформление_V.2
                    {
                      let sheetName = name.getName();
                      if(sheetName.indexOf("Оформление_V.2")>=0) // Заходим в лист с таким именем
                      {
                        let data = sheet.getSheetByName(sheetName);
                        let tempRange = data.getRange(2, 14, data.getLastRow()-1, 4).getValues()
                        
                        for(let i = 0; i<tempRange.length; i++)
                        {
                          for(let j = 0; j<temporaryArray.length; j++)
                          {
                            if(tempRange[i][2] === temporaryArray[j][2])
                            {
                              if(tempRange[i][0] != temporaryArray[j][0])
                              {
                                Logger.log(tempRange[i][2] + " <> " + temporaryArray[j][2]);
                                Logger.log(tempRange[i][0] + " <> " + temporaryArray[j][0]);
                                Logger.log(i+2);
                                Logger.log(`https://docs.google.com/spreadsheets/d/${item}`);
                                Logger.log("===============================================================================");
                                data.getRange(i+2, 14).setValue(temporaryArray[j][0]); 
                              }
                              if(tempRange[i][1] != temporaryArray[j][1])
                              {
                                Logger.log(tempRange[i][2] + " <> " + temporaryArray[j][2]);
                                Logger.log(tempRange[i][1] + " <> " + temporaryArray[j][1]);
                                Logger.log(i+2);
                                Logger.log(`https://docs.google.com/spreadsheets/d/${item}`);
                                Logger.log("===============================================================================");                                
                                data.getRange(i+2, 15).setValue(temporaryArray[j][1]);                                
                              }
                            }  
                          }
                        }     
                      }  
                    });
  });
}


//==========================================================================================================================================
//=============================Создание UUID для строк с датой по столбцу Т=================================================================
//==========================================================================================================================================
function createUUID() {  
  mainSheet.getRange(2, 20, mainSheet.getLastRow()-1, 2).getValues() // Берем список всех значений по столбцам T и U
                                                        .map((row, i) => row.concat((i+2))) // Вкладываем порядковый номер строки в каждый массив в массиве.
                                                        .filter(row => row[0] && !row[1]) // фильтруем массив от пустых значений по столбцу T и свободных U. Т.е. нам нужны строки только где стоит одна лишь дата
                                                        .forEach(item => mainSheet.getRange(item[2], 21).setValue(Utilities.getUuid())); // Для строк в которых столбец Т проставлен, а столбец U пустой, проставляем в U UUID
}

//==========================================================================================================================================
//===========Перенос строк из общей ЛК в главные таблицы, для строк у которых указана новая дата по столбцу Т и сгенерирован UUID===========
//==========================================================================================================================================
function transferToMain(){
  let transferlist = mainSheet.getRange(2, 1, mainSheet.getLastRow()-1, mainSheet.getLastColumn()).getValues()
                                                                                                  .map((row, i) => row.concat((i+2)))
                                                                                                  .filter(row =>row[20]);
  sheetsIds.forEach(id=>{                          
                              let tempList = transferlist.filter(row=>row[16]==id);              
                              let main = SpreadsheetApp.openById(id);
                              main.getSheets().filter(sheet=>sheet.getName()
                                              .indexOf("Оформление_V.2")>=0)
                                              .map(sheet=>sheet.getName())
                                              .forEach(sheetName=>{
                                                                    let sheet = main.getSheetByName(sheetName);
                                                                    let datesColumn = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues().map(date=> date[0] ? date[0].toDate() : "");
                                                                    let uuids = sheet.getRange(2, 16, sheet.getLastRow()-1, 1).getValues().filter(cell=>cell[0]).map(cell=>cell[0]);                                                                   
                                                                    let sheetId = sheet.getSheetId().toFixed(10).replace(/\.?0+$/,'');
                                                                    tempList.filter(row=>row[17]==sheetId&&uuids.indexOf(row[20])<0)
                                                                            .forEach(row=>{
                                                                                            let newRow = row.slice(0, 15);
                                                                                            newRow[1] = row[19];
                                                                                            newRow[15] = row[20];
                                                                                            let obj = findByDatePosition(row[19], datesColumn);
                                                                                            let pos = obj[0];
                                                                                            datesColumn = obj[1];
                                                                                            sheet.insertRowAfter(pos);
                                                                                            pos++;
                                                                                            Logger.log(newRow);
                                                                                            sheet.getRange(pos, 1, 1, newRow.length).setValues([newRow]);
                                                                                          });
                                                                  });                          
                            });
}

//==========================================================================================================================================
//=====================================Вспомогательная функция поиска позиции по дате из колонки дат========================================
//==========================================================================================================================================
function findByDatePosition(date, datesColumn){
  date = date.toDate();
  let pos = 0;
  datesColumn.indexOf(date) != -1 ? pos=datesColumn.indexOf(date)+2:pos=-1; 
  pos != -1 ? datesColumn.splice(pos, 0, date) : pos = datesColumn.last()+2;  
  return[pos, datesColumn];
}
















