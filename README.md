# 在 Google Spreadsheet 中擷取台灣股市上市/上櫃股票價格

在  Google SPreadsheet，透過巨集熱鍵，擷取台灣證券交易所之股價資料

參考 Asoul/tsrtc 在台灣股票即時爬蟲 Taiwan Stock Exchange Real Time Crawler 撰寫的 getStockInfo.jsp 的分析資料

## 參考資料

* [用Google試算表追蹤百檔以上台股(上市/上櫃/興櫃)股價](https://wiki0918.pixnet.net/blog/post/222332253-用google試算表取得台股%28上市-上櫃%29股票報價)
* [Asoul/tsrtc 的證交所 API Document （偽）](https://github.com/Asoul/tsrtc)
* [Google 參考文件](https://developers.google.com/apps-script/reference/spreadsheet)

## 使用方法

* 直接開啟Google Spreadsheet [股票觀察_熱鍵分享板](https://docs.google.com/spreadsheets/d/1K0OgjeL3uMZZ7JvbjE-MbkbKldgjjPEt9DYyzuuuelU/edit?usp=sharing)的分享檔案
* 按 Ctrl + Alt+ Shift + 1 即可執行執行
* 若無法從台灣證券交易所獲得資料，在 M1 儲存格會留下查詢的字串，M2 會顯示重新執行的熱鍵
* 按 Ctrl + Alt + Shift + 0 再次向台灣證券交易所，索取資料
(請注意，TWSE 有 request limit, 每 5 秒鐘 3 個 request，超過的話會被 ban 掉，請自行注意)

## 建立個人的觀察檔案

* 直接開啟Google Spreadsheet [股票觀察_熱鍵分享板](https://docs.google.com/spreadsheets/d/1K0OgjeL3uMZZ7JvbjE-MbkbKldgjjPEt9DYyzuuuelU/edit?usp=sharing)的分享檔案
* 點選檔案=>建立複本，即可以建立個人專屬的檔案

## 異常處理

檢查 Google Spreadsheet 的巨集設定，確定 setStocksPrice, retry_setStocksPrice 被正確載入，並完成了熱鍵的設定

* setStocksPrice 的熱鍵為 Ctrl+Alt+Shift+1
* retry_setStocksPrice 的熱鍵為 Ctrl+Alt+Shift+0

若函式被載入，請透過 工具=>巨集=>匯入 ，將上述兩個函數匯入。並透過 工具=>巨集=>管理巨集 完成熱鍵的設定。

## 函數說明

共 4 個函數，分別為 setStocksPrice()、retry_setStocksPrice()、getStockQueryString() 、getStockPrice(stocksID)

* setStocksPrice() 主要的程式，呼叫 getStockQueryString() 得到查詢字串，呼叫 getStockPrice(stocksID) 取得包含股票資訊的 json 檔，並完成價格的更新
* retry_setStocksPrice() 重試的程式，呼叫 getStockPrice(stocksID) 取得包含股票資訊的 json 檔並完成價格的更新
* getStockQueryString() 用來組成「股票查詢字串」的函數，由表格的 A4 欄位往下取得，D欄 狀態設定為 1 的資料。 A4 往下至第一個空白的 A 欄位即停止
* getStockPrice(stocksID) 向查詢台灣證交所的 getStockInfo.jsp 查詢，並得到 json 資料回傳

### setStocksPrice
```javascript
function setStocksPrice() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('M1').setValue('');
  spreadsheet.getRange('M2').setValue(""); 
  var StockQueryString= getStockQueryString();
  //spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('M1').setValue(StockQueryString);
  
  var stockPriceDetail;
  
  // 接取資料測試，當傳回錯誤時，在 M2 的位置提供重新測試的熱鍵
  try {
    stockPriceDetail=getStockPrice(StockQueryString);
    spreadsheet.getRange('M2').setValue('');
    spreadsheet.getRange('M1').setValue('')
    var j=0;
    spreadsheet.getRange('C4').activate();
    for(i=0;i<stockPriceDetail.length;i++)
    {
      while( spreadsheet.getCurrentCell().offset(i+j,1).getValue() != 1)
      {
        j++; 
      }
      // obj.n 為股票的名稱   
      // spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].n);
      // obj.c 為股票代號
      // spreadsheet.getCurrentCell().offset(i+j,2).setValue(stockPriceDetail[i].c);
      
      // obj.z 為最近的成交價
      spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].z);   
      ;   
    }}
  catch(err){
    spreadsheet.getRange('M2').setValue("Ctrl + Alt + Shift + 0 to Retry");  
  }
 
};
```
### retry_setStocksPrice
```javascript
function retry_setStocksPrice() {
// 當無法取得資料來，用來再重新查詢使用的
  var spreadsheet = SpreadsheetApp.getActive();
  var StockQueryString=spreadsheet.getRange('M1').getValue();
  
  var j=0;
  
  try{
    spreadsheet.getRange('M2').setValue("重新抓取資料中.....");      
    var stockPriceDetail=getStockPrice(StockQueryString);
    spreadsheet.getRange('C4').activate();
    for(i=0;i<stockPriceDetail.length;i++)
    {
      while( spreadsheet.getCurrentCell().offset(i+j,1).getValue() != 1)
      {
        j++; 
      }
      // obj.n 為股票的名稱   
      // spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].n);
      // obj.c 為股票代號
      // spreadsheet.getCurrentCell().offset(i+j,2).setValue(stockPriceDetail[i].c);
      
      // obj.z 為最近的成交價
      spreadsheet.getCurrentCell().offset(i+j,0).setValue(stockPriceDetail[i].z);   
      
    }
    spreadsheet.getRange('M2').setValue("");
    spreadsheet.getRange('M1').setValue('');
  } 
  catch(err)
  {
    spreadsheet.getRange('M2').setValue("Ctrl + Alt + Shift + 0 to Retry"); 
  }
};
```
###  getStockQueryString
```javascript
function getStockQueryString() {
// 這是一個用來組成「股票查詢字串」的巨集
  var spreadsheet = SpreadsheetApp.getActive();
// 股票的代碼由 A4 開始，上市股請的代號為  tse_0000 ；上櫃為 otc_0000  
  var currentCell = spreadsheet.getRange('A4').activate();
  var stock_IDs=spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).getValues();
  var stock_ID_inquery="";
  
  for (i=0;i<stock_IDs.length;i++)
  {
    if (i>0)
    {
// 股票未賣出，為持有時請設為 1 
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
```
### getStockPrice()
```javascript
function getStockPrice(stocksID)
{
  // 連結至台灣證卷交易所的  getStockInfo.jsp 的 API ，取回查詢股票資訊的 jsop 格式資料
  //stocksID="otc_5356.tw%7cotc_8069.tw%7ctse_1532.tw"; //|otc_8064.tw|otc_3586.tw";
  //var host_address="117.56.218.179"; 
  var host_address="mis.twse.com.tw";
  var url = "http://"+host_address+"/stock/api/getStockInfo.jsp?json=1&delay=0&ex_ch="+stocksID+"&_="+Date.now();
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText("UTF-8");
  var data = JSON.parse(json);
  return data.msgArray;   
};
```


