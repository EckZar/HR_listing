Array.prototype.last = function(){ // Поиск последнего не пустого элемента в массиве
  for(let i = this.length; i>0; i--)
  {
    if(this[i])
    {
      return i;
    }
  }
}

Object.prototype.toDate = function(){ // Преобразование даты в формат dd.mm.yyyy
  let date = new Date(this);
  return date.getDate() + "." + (date.getMonth()+1) + "." + date.getFullYear();
}
