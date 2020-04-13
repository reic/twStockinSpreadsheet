function FastMode_construct(){
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getSheets()[0];

  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A3').setBackground("#cfe2f3").setValue("執行中....");
// 通過指定一個特定的 Cell ，然後運用 getDataRegion (輔以 SpreadsheetApp.Dimension.ROWS 指定方向
// 再用 getA1Notation 取得  A4:A18 的資料範圍
  var range=sheet.getRange("a6").getDataRegion(SpreadsheetApp.Dimension.ROWS).getA1Notation();

// 取得 getRange 取得 Cells 的值，為二維陣列 array[][]，即使取一欄也為  [[a4],[a5],[a6]]
  var myarray=sheet.getRange(range.replace(/:A/,":D")).getValues();
  
// stocks 為股票代碼矩陣
// price 為股票價格矩陣
  var stocks=[];
  var price=[];
  var unistock=[];
  var querystring='';
  var stockPriceDetail;

  clearStocksPrice();
//設定資訊
  [price,querystring]=preSetQuery(myarray,range,stocks,price,querystring);
  
  
  var url = "http://mis.twse.com.tw/stock/api/getStockInfo.jsp?json=1&delay=0&ex_ch="+querystring+"&_="+Date.now();
  sheet.getRange('N1').setValue(url); 
  
  try{  
// 取得股票的資料
    stockPriceDetail=getStockPrice(url);
    sheet.getRange('N1').setValue('');
// 透過 function 設定變數
    setNewStockPrice(myarray,price,range,stockPriceDetail);
    sheet.getRange('A3').setBackground("#d9ead3").setValue("完成更新");
  }catch(err)
  {         
//自動重新嘗試連線一次  
    
// Google Utilities API    
    Utilities.sleep(5 * 1000);
    try{
      stockPriceDetail=getStockPrice(url);
      sheet.getRange('N1').setValue('');
      setNewStockPrice(myarray,price,range,stockPriceDetail);
      sheet.getRange('A3').setBackground("#d9ead3").setValue("完成更新");      
    }catch(err){
      Logger.log(err);
      sheet.getRange('A3').setBackground("#fce5cd").setValue("請重試");
    }   
  } 
}

  

function clearStocksPrice() {
// 慢慢的清資料
  var spreadsheet = SpreadsheetApp.getActive();
  var currentCell = spreadsheet.getRange('C5').activate();
  var stock_IDs=spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).getValues();
  var stock_ID_inquery='';
  
  for (i=0;i<stock_IDs.length;i++)
  {
      if(currentCell.offset(i,1).getValue() ==1)
      {
        currentCell.offset(i,0).setValue("").offset(i,2).setValue("");        
      }   
   }
     
  };


function preSetQuery(myarray,range,stocks,price, querystring)
{
var sheet = SpreadsheetApp.getActiveSheet();  
// 股票為持有狀態，才會更新
// myarray 用 foreach 的 function callback 方式 更新資料 myarray 每一列的資料
// stock 代碼只取得需要更新的代碼
  myarray.forEach(function (value) {      
    if(value[3]==1) { price.push([""]); stocks.push(value[0]);    }
    else  { price.push([value[2]]);  }});  
// 查詢前，先調整現價的 欄位內容
  sheet.getRange(range.replace(/A/gi,"C")).setValues(price);   
// 完成股票代碼查詢
// 消除 stocks 內，相同的股票代碼  
// ex: ["5356", "00677U", "8069", "8069", "1532", "8069"] => ["5356", "00677U", "8069", "1532"]
  var unistock=stocks.filter(function(element, index, arr){return arr.indexOf(element) == index; });
// 建構查詢字串  
  unistock.forEach(function(value){   
    if(querystring.length >0) {querystring+="%7c";}
       querystring+="tse_"+value+".tw%7cotc_"+value+".tw";  });
// function 只能回傳一個值，當有多值需要回傳的時候，可以透過包裝成矩陣物件的形式
// 再透過矩陣接收變數
  return [price,querystring];  
};



function getStockPrice(url)
{
  // 連結至台灣證卷交易所的  getStockInfo.jsp 的 API ，取回查詢股票資訊的 jsop 格式資料
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText("UTF-8");
  var data = JSON.parse(json);
  return data.msgArray;  
};


function setNewStockPrice(myarray,price,range,stockPriceDetail){
var sheet = SpreadsheetApp.getActiveSheet();  
var updown=[];
  for(i=0;i<myarray.length;i++)
  {
    if(myarray[i][3]==1){
      for( item of stockPriceDetail)
      {
         if(myarray[i][0]==item.tk1.split(".",1)) 
         {
           if (item.z !='-') {
			   price[i][0]=item.z; 
			   updown.push([(item.z-item.y).toString()]);}
           else {
                 if(item.b !='-')  { price[i][0]=item.b.split("_",1); 
				   updown.push([(item.b.split("_",1)-item.y).toString()]);}
                 else{ price[i][0]=item.y; updown.push(["-"]);}
               }       
        }
   }}else{
      updown.push(['-']);
   }
}
   sheet.getRange(range.replace(/A/gi,"C")).setValues(price); 
   sheet.getRange(range.replace(/A/gi,"E")).setValues(updown); 
};
