function fetchLatestValue(brand_code){
  const get_url = "https://kabutan.jp/stock/kabuka?code="+String(brand_code);
  const latest_value_index = 3;
  
  var stock0_entire_element;
  var stock0_value_table;
  var stock0_closing_price;
//  try{
    /* [TBD] Shall not read value at holiday */
    var html = UrlFetchApp.fetch(get_url).getContentText('UTF-8');
    stock0_entire_element = Parser.data(html).from('<table class="stock_kabuka0">').to('</table>').build();
    stock0_value_table = Parser.data(stock0_entire_element).from('<td>').to('</td>').iterate();
    stock0_closing_price = Number(stock0_value_table[latest_value_index].replace(',',''));
//    Logger.log(stock0_closing_price);
    
    if(stock0_closing_price >= 0){
      return stock0_closing_price;
    }
//  }catch(e){
    /* Brand Value Fetch Error Statement */
//  }
  return -1;
}

function fetchBrandCodeList(sheet){
  var brand_code_list = sheet.getRange('A:A').getValues();

  return brand_code_list.filter(isInteger); /* [TBD] Should fix range limitation */
}

function writeTodaysValue(sheet, brand_value, row_index, todays_column, start_row){
  if(todays_column >= 0){
    sheet.getRange(start_row + row_index, todays_column).setValue(brand_value);
  }else{
    sheet.getRange(start_row + row_index, todays_column).setValue("SomeThing Error");
  }
}

function searchTodaysColumnValue(sheet){
  const date_row = 2;
  const date_list = sheet.getRange('2:2').getValues(); /* [TBD] 2dim Array*/
  
  var date = new Date();
  var today = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
  for(var i = 1; i < date_list[0].length; i++){
    var target_date;
    
    try{
      target_date = Utilities.formatDate(date_list[0][i], 'Asia/Tokyo', 'yyyyMMdd');
    }catch(e){
      /* Fetch Date Error Statement */
    }
    
    
    if(target_date == today){
      return i+i;
      break;
   }
  }
  return -1;
}

function isInteger(entry){
  if(entry > 0){
    return true;
  }
  return false;
}

/* Main */
function myFunction() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet(); 
  
  const todays_column = searchTodaysColumnValue(sheet);
  var brand_code_list = fetchBrandCodeList(sheet);
  
  const start_row = 3;
 
  var write_index = 0;
  for(var i = 0; i < brand_code_list.length; i++) {
    var entry_latest_value = fetchLatestValue(brand_code_list[i]);
    if(entry_latest_value >=0){
      writeTodaysValue(sheet, entry_latest_value, i, todays_column, start_row);
    }
    Utilities.sleep(850);
  }
}

function eventFunction(sheet) {
  const todays_column = searchTodaysColumnValue(sheet);
  var brand_code_list = fetchBrandCodeList(sheet);
//  Browser.msgBox(brand_code_list);
  
  const start_row = 3;
 
  var write_index = 0;
  for(var i = 0; i < brand_code_list.length; i++) {
    var entry_latest_value = fetchLatestValue(brand_code_list[i]);
    if(entry_latest_value >=0){
      writeTodaysValue(sheet, entry_latest_value, i, todays_column, start_row);
    }
    Utilities.sleep(850);
  }
}

/* Event Trigger */
function onChange(event) {
  var sheet = event.source.getActiveSheet();
  var cell = event.source.getActiveRange();

  if (sheet.getName() == "シート1" && cell.getColumn() == 1) {
    eventFunction(sheet);
    Logger.log("Name loop Enterd:");
    const todays_column = searchTodaysColumnValue(sheet);
    
    try{
      if (Number(event.value) >= 0) {
        Logger.log("Event Value Enterd:");
        var entry_latest_value = fetchLatestValue(event.value);
        Browser.msgBox("こんにちは"+entry_latest_value);
        Logger.log("Event Value Enterd:"+entry_latest_value);
        writeTodaysValue(sheet, entry_latest_value, event.range.getRow(), todays_column, 0);
      }else{
        writeTodaysValue(sheet, entry_latest_value, event.range.getRow(), -1, 0);
      }
    }catch(e){
      /* Invalid Value Entered */
    }
  }
}