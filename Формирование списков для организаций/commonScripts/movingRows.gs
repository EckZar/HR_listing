//==========================================================================================================================================
//===========Перенос строк из общей ЛК в главные таблицы, для строк у которых указана новая дата по столбцу Т и сгенерирован UUID===========
//==========================================================================================================================================
function transferToMain(){
  let transferlist = lkSheet.getRange(2, 1, lkSheet.getLastRow()-1, lkSheet.getLastColumn()).getValues()
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
                                                                    let datesColumn = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues().map(date=> {
                                                                                                                                                              let tempDate;
                                                                                                                                                              if(date[0])
                                                                                                                                                              {
                                                                                                                                                                tempDate=new Date(date[0]);
                                                                                                                                                                tempDate = tempDate.getDate() + "." + (tempDate.getMonth()+1) + "." + tempDate.getFullYear();
                                                                                                                                                              }
                                                                                                                                                              else {tempDate = ""}
                                                                                                                                                              return tempDate;
                                                                                                                                                           });      (datesColumn)                                                               
                                                                    let uuids = sheet.getRange(1, 16, sheet.getLastRow()-1, 1).getValues().map(cell=>cell[0]);                                                                
                                                                    let sheetId = sheet.getSheetId().toFixed(10).replace(/\.?0+$/,'');
                                                                    tempList.filter(row=>row[17]==sheetId&&uuids.indexOf(row[20])<0)
                                                                            .forEach(row=>{
                                                                                            lkSheet.getRange(row[21], 14).setValue("принят");
                                                                                            sheet.getRange(uuids.indexOf(row[15])+1, 14).setValue("принят");
                                                                                            let newRow = row.slice(0, 15);
                                                                                            newRow[1] = row[19];
                                                                                            newRow[15] = row[20];
                                                                                            let obj = findByDatePosition(row[19], datesColumn);
                                                                                            let pos = obj[0];
                                                                                            datesColumn = obj[1];
                                                                                            sheet.insertRowAfter(pos);
                                                                                            pos++;
                                                                                            Logger.log(pos);
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
  let tempDate = new Date(date);
  tempDate = tempDate.getDate() + "." + (tempDate.getMonth()+1) + "." + tempDate.getFullYear();
  let pos = 0;

  datesColumn.indexOf(tempDate) != -1 ? pos=datesColumn.indexOf(tempDate)+2:pos=-1; 

  let lastPos;

  for(let i = datesColumn.length; i>0; i--)
  {
    if(datesColumn[i])
    {
      lastPos = i;
      break;
    }
  }

  pos != -1 ? datesColumn.splice(pos, 0, tempDate) : pos = lastPos+2;  
  return[pos, datesColumn];
}

//==================================================================================================
//================================фуникция сортировки по дате=======================================
//==================================================================================================
function checkDatesQueue(){
  let dates = lkSheet.getRange("B2:B").getValues().filter( date => date[0] != "").map( date => date[0]);
  for(let i = dates.length; i>0; i--)
  {
    if(dates[i] < dates[i-1])
    {
      Logger.log("Сортируем не последовательность");
      lkSheet.getRange(2, 1, lkSheet.getLastRow()-1, lkSheet.getLastColumn()).sort({column: 2, ascending: true});
      break;
    }
  }
}
