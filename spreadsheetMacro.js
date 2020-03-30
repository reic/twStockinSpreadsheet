function setStocksPrice() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A3').setValue("執行中....");
// Google SpreadSheet 
// M1 會做為 查詢股票資料暫存使用
// M2 為當發生無法正確取得 getStockInfo.jsp 回傳資料時，會提示重新執行的熱鍵
  spreadsheet.getRange('M1').setValue('');
  spreadsheet.getRange('M2').setValue(''); 

// 呼叫 getStockQueryString() 函式，用來串接 getStockInfo.jsp 的 ex_ch 的上市櫃股票查詢變數
  var StockQueryString= getStockQueryString();

// 暫存查詢字串在 M1 儲存格
  spreadsheet.getRange('M1').setValue(StockQueryString);
  
  var stockPriceDetail;
  
  // 接取資料測試，當傳回錯誤時，在 M2 的位置提供重新測試的熱鍵  
  try {
    stockPriceDetail=getStockPrice(StockQueryString);
    spreadsheet.getRange('M2').setValue('');
    spreadsheet.getRange('M1').setValue('')
    
 // 取得的資料寫入 google spreadsheet
    var j=0;
    spreadsheet.getRange('C4').activate();
    for(i=0;i<stockPriceDetail.length;i++)
    {
      while( spreadsheet.getCurrentCell().offset(i+j,1).getValue() != 1)
      {
        j++; 
      }          
// obj.z 為最近的成交價 obj.y 昨日收盤價
// obj.y 昨日收盤價
// obj.b 買入 5 檔，取最高的買入價
      if (stockPriceDetail[i].z =="-")
      {
        if(spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].b =="-"))
        {
             spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].y);
        }else{
             spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].b.split("_",1));  
        }
      } else
      {
        spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].z);   
      }   
    }
    spreadsheet.getRange('A3').setValue("完成更新");
  }
  catch(err){
// 無法取得資料的例外處理
// 在 M2 儲存格提示重新執行的熱鍵
    spreadsheet.getRange('M2').setValue("Ctrl + Alt + Shift + 0 to Retry");  
    spreadsheet.getRange('A3').setValue("請重試");
  }
 
};


function retry_setStocksPrice() {
// 當無法取得資料來，用來再重新查詢使用的
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A3').setValue("執行中....");
  
// 由 M1 儲存格取得 getStockInfo.jsp 的查詢字串
  var StockQueryString=spreadsheet.getRange('M1').getValue();
  spreadsheet.getRange('M2').setValue("重新抓取資料中....."); 
  
  try{
         
    var stockPriceDetail=getStockPrice(StockQueryString);
 
 // 取得的資料寫入 google spreadsheet
    spreadsheet.getRange('C4').activate();
    var j=0;
    for(i=0;i<stockPriceDetail.length;i++)
    {
      while( spreadsheet.getCurrentCell().offset(i+j,1).getValue() != 1)
      {
        j++; 
      }
// obj.z 為最近的成交價 obj.y 昨日收盤價
// obj.y 昨日收盤價
// obj.b 買入 5 檔，取最高的買入價
      if (stockPriceDetail[i].z =="-")
      {
        if(spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].b =="-"))
        {
             spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].y);
        }else{
             spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].b.split("_",1));  
        }
      } else
      {
        spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].z);   
      }    
    }
// 重置 M1, M2 儲存格
    spreadsheet.getRange('M2').setValue('');
    spreadsheet.getRange('M1').setValue('');
    spreadsheet.getRange('A3').setValue("完成更新");
  } 
  catch(err)
  {
    spreadsheet.getRange('M2').setValue("Ctrl + Alt + Shift + 0 to Retry"); 
    spreadsheet.getRange('A3').setValue("請重試");
  }
};


function getStockQueryString() {
// 這是一個用來組成「股票查詢字串」的巨集
  var spreadsheet = SpreadsheetApp.getActive();
// 股票的代碼由 A4 開始，上市股請的代號為  tse_0000 ；上櫃為 otc_0000  
// 取得連續的資料
  var currentCell = spreadsheet.getRange('A4').activate();
  var stock_IDs=spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).getValues();
  var stock_ID_inquery='';
  
  for (i=0;i<stock_IDs.length;i++)
  {
    if (i>0)
    {
// 仍持有股票時，會在  D 欄設定變數為 1
// 僅查詢仍持有股票的部分
// Reic 懶惰，為了簡便處理，同一檔股票若分為兩列，會向 getStockInfo.jsp 查詢兩次
      if(parseInt(currentCell.offset(i,3).getValue(),10) ==1)
      {
        currentCell.offset(i,2).setValue("");
        stock_ID_inquery+="%7c"+stock_IDs[i]+".tw";
      }
    }else{
      if(parseInt(currentCell.offset(i,3).getValue(),10)==1)
      {
        currentCell.offset(i,2).setValue("");
        stock_ID_inquery+=stock_IDs[i]+".tw";
      }
    }
   }
   return stock_ID_inquery;  
};


function getStockPrice(stocksID)
{
  // 連結至台灣證卷交易所的  getStockInfo.jsp 的 API ，取回查詢股票資訊的 jsop 格式資料
  var host_address="mis.twse.com.tw";
  var url = "http://"+host_address+"/stock/api/getStockInfo.jsp?json=1&delay=0&ex_ch="+stocksID+"&_="+Date.now();
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText("UTF-8");
  var data = JSON.parse(json);
  return data.msgArray;  
};
