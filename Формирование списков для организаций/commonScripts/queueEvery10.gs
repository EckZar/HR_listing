function ten() {

  if(getPropStatus("insertNewRows")==1)
  {
    Logger.log("updateSheet() function still working");
    return;
  } 

  setProp("insertNewRows", 1);

  try{
    let lock = LockService.getScriptLock();
    lock.tryLock(5000);
    bulkUpdateSheet();

    lock.tryLock(5000);    
    rowsSyncronisation();

    lock.tryLock(5000);
    removeUnrecognizedUUID();

    lock.tryLock(5000);
    deleteIsiRows();

    lock.tryLock(5000);
    removeDuplicatesByUUID();

    lock.tryLock(5000);
    deleteEmptyRow();
    
    lock.tryLock(5000);
    checkDatesQueue();

    lock.tryLock(5000);
    transferToMain();
  }
  catch(e){
    Logger.log(e);
  }
  finally{
    setProp("insertNewRows", 0);
  }

}
