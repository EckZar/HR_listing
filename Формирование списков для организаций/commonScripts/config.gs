const regionsSheets = {
   "msk": {
     "main":"-",
     "isi": ""
   },
   "spb": {
     "main":"",
     "isi": "-"
   },
   "krd": {
     "main":"",
     "isi": "-"
   },
   "ptr": { // Потребность
     "main":"-",
     "isi": "-"
   },
   "pvl": { // Поволжье
     "main":"-",
     "isi": "-"
   },
   "yar":{ // Ярославль
     "main": "-U",
     "isi": ""
   },
   "kl":{ // Калуга
     "main": "-",
     "isi": "-"
   },
   "rz":{ // Рязань
     "main": "",
     "isi": "-"
   }
}

const lkSheet = SpreadsheetApp.openById("-").getSheetByName("Оформление_V.2");

const sheetsIds = [

];

const enterString = String.fromCharCode(10);
