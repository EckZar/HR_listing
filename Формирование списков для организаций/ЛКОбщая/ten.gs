function ten() {

  if(getPropStatus("insertNewRows")==1)
  {
    Logger.log("updateSheet() function still working");
    return;
  } 

  setProp("insertNewRows", 1);

  try
  {
    transferToMain();
  }
  catch(e){}
}
