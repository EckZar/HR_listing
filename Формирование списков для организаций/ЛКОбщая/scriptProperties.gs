function getPropStatus(propName) {  
  let props = PropertiesService.getScriptProperties().getProperty(propName);
  Logger.log(props);
  return props;
}

function setProp(propName, status){
  PropertiesService.getScriptProperties().setProperty(propName, status);
}

function dropStatus(){
  setProp("insertNewRows", 0);
}
