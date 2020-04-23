

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
  
  
//  var url = "http://mis.twse.com.tw/stock/api/getStockInfo.jsp?json=1&delay=0&ex_ch="+querystring+"&_="+Date.now();
// 因為 mis.twse.com.tw 有許多 ip，因此導入 NSLookup 每次查詢前，先找到適合的 IP
  var hostname="mis.twse.com.tw";
  var url = "http://mis.twse.com.tw/stock/api/getStockInfo.jsp?json=1&delay=0&ex_ch="+querystring+"&_="+Date.now();

  sheet.getRange('N1').setValue(url); 
  
  try{  
// 取得股票的資料
// url.replace 即是將 hostname 轉換為 ip  的過程
    stockPriceDetail=getStockPrice(url.replace(hostname,NSLookup(hostname)));
    sheet.getRange('N1').setValue('');
// 透過 function 設定變數
    setNewStockPrice(myarray,price,range,stockPriceDetail);
    sheet.getRange('A3').setBackground("#d9ead3").setValue("完成更新");
  }catch(err)
  {         
//自動重新嘗試連線一次  
    
// Google Utilities API    
    Utilities.sleep(3 * 1000);
    try{
// url.replace 即是將 hostname 轉換為 ip  的過程     
      stockPriceDetail=getStockPrice(url.replace(hostname,NSLookup(hostname)));
      sheet.getRange('N1').setValue('');
      setNewStockPrice(myarray,price,range,stockPriceDetail);
      sheet.getRange('A3').setBackground("#d9ead3").setValue("完成更新");      
    }catch(err){
      Logger.log(err);
      sheet.getRange('A3').setBackground("#fce5cd").setValue("請重試");
    }
   
  }
 
}

function re_Construct(){
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A3').setBackground("#cfe2f3").setValue("執行中....");
  var hostname="mis.twse.com.tw";
// 由資料表取得 url String
  var url=sheet.getRange('N1').getValue();
 

  if(url.length<3){
    sheet.getRange('A3').setBackground("#fce5cd").setValue("請點選執行");
    return;    
  };
  
// 通過指定一個特定的 Cell ，然後運用 getDataRegion (輔以 SpreadsheetApp.Dimension.ROWS 指定方向
// 再用 getA1Notation 取得  A4:A18 的資料範圍
  var range=sheet.getRange("a6").getDataRegion(SpreadsheetApp.Dimension.ROWS).getA1Notation();

// 取得 getRange 取得 Cells 的值，為二維陣列 array[][]，即使取一欄也為  [[a4],[a5],[a6]]
  var myarray=sheet.getRange(range.replace(/:A/,":D")).getValues();  

// price 為股票價格矩陣
  var price=[];
  var stockPriceDetail;  
  myarray.forEach(function (value){price.push([value[2]]);});
  
  try{  
// 取得股票的資料
// url.replace 即是將 hostname 轉換為 ip  的過程
    stockPriceDetail=getStockPrice(url.replace(hostname,NSLookup(hostname)));
    sheet.getRange('N1').setValue('');    
// 透過 function 設定變數
    setNewStockPrice(myarray,price,range,stockPriceDetail);
    sheet.getRange('A3').setBackground("#d9ead3").setValue("完成更新");
  }catch(err)
  {
//自動重新嘗試連線一次    
// google Utilities API， 3 秒後再重試    
    Utilities.sleep(3 * 1000);
    try{
// url.replace 即是將 hostname 轉換為 ip  的過程      
      stockPriceDetail=getStockPrice(url.replace(hostname,NSLookup(hostname)));
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
  
  for (i=0;i<stock_IDs.length;i++)
  {
      if(currentCell.offset(i,1).getValue() ==1)
      {
        currentCell.offset(i,0).setValue("").offset(0,2).setValue("");        
      }   
   }
 
  };


function preSetQuery(myarray,range,stocks,price, querystring)
{
var sheet = SpreadsheetApp.getActiveSheet();  
// 股票為持有狀態，才會更新
// myarray 用 foreach 的 function callback 方式 更新資料 myarray 每一列的資料
// stock 代碼只取得需要更新的代碼
// 0417 增加迴圈的效率, remove if --else , use the rule : Less Nesting, Return Early in foreach
// 0417 判斷是否為 1，進行資料處理，返回執行下一個迴圈。 可以節省一次 else 的判斷
  myarray.forEach(function (value) {      
    if(value[3]==1) { price.push([""]); stocks.push(value[0]); return;}
     price.push([value[2]]);  });  

// 查詢前，先調整現價的 欄位內容
//  目前由 clearStockPrice 取代，增加趣味性
//  sheet.getRange(range.replace(/A/gi,"C")).setValues(price);   
//  sheet.getRange(range.replace(/A/gi,"E")).setValues(price);   
  
// 完成股票代碼查詢
// 消除 stocks 內，相同的股票代碼  
// ex: ["5356", "00677U", "8069", "8069", "1532", "8069"] => ["5356", "00677U", "8069", "1532"]
  var unistock=stocks.filter(function(element, index, arr){return arr.indexOf(element) == index; });
// 建構查詢字串  
// 增加 foreach( callback(value, index) , 將 querystring.length 判斷換為 index 判斷
  unistock.forEach(function(value,index){   
    if(index!=0) {querystring+="%7c";}
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
  // 修正回傳第一個字元偶會出現 < 的問題
  if (json.startsWith("<")) json=json.replace(/</,'');
  var data = JSON.parse(json);
  return data.msgArray;  
};


function setNewStockPrice(myarray,price,range,stockPriceDetail){
  var sheet = SpreadsheetApp.getActiveSheet();
  var updown = [];
// 透過 map 直接整理取回來的 股票資訊，若無  z 即時資訊，則取 買方 5 檔的最高價
  var stockset = stockPriceDetail.map(function (item) { item.tk0 = item.tk0.split(".")[0]; item.z = item.z != "-" ? item.z : item.b.split("_")[0]; return { id: item.tk0, z: item.z, y: item.y } });
  myarray.forEach(function (item, index) {
    if (item[3] != 1) { updown.push(['-']); return; }
    id = item[0]; // 取到表單上的股票代碼
// 透過 filter 找到股票的 現價、昨天
    [{ z, y }] = stockset.filter(function (item) { return item.id == id });
    price[index][0] = z;
    updown.push([z - y]);
  })
  sheet.getRange(range.replace(/A/gi, "C")).setValues(price);
  sheet.getRange(range.replace(/A/gi, "E")).setValues(updown);
}



//   myarray.forEach(function (item, index) {
//     if (item[3] != 1) { updown.push(['-']); return; }
//     for (stock of stockPriceDetail) {
//       if (item[0] != stock.tk1.split(".", 1)) { continue; }
//       if (stock.z != '-') { price[index][0] = stock.z; updown.push([(stock.z - stock.y)]); return; }
//       if (stock.b != '-') { price[index][0] = stock.b.split("_"); updown.push([(price[index][0] - stock.y)]); return; }
//       price[index][0] = stock.y; updown.push(["-"]);
//       break;
//     }
//   })
//   sheet.getRange(range.replace(/A/gi, "C")).setValues(price);
//   sheet.getRange(range.replace(/A/gi, "E")).setValues(updown);
// };



//   for(i=0;i<myarray.length;i++)
//   {
//     if(myarray[i][3]==1){
//       for( item of stockPriceDetail)
//       {
//          if(myarray[i][0]==item.tk1.split(".",1)) 
//          {
//            if (item.z !='-') {
// 			   price[i][0]=item.z; 
// 			   updown.push([(item.z-item.y).toString()]);}
//            else {
//                  if(item.b !='-')  { price[i][0]=item.b.split("_",1); 
// 				   updown.push([(item.b.split("_",1)-item.y).toString()]);}
//                  else{ price[i][0]=item.y; updown.push(["-"]);}
//                }       
//         }
//    }}else{
//       updown.push(['-']);
//    }
// }
//    sheet.getRange(range.replace(/A/gi,"C")).setValues(price); 
//    sheet.getRange(range.replace(/A/gi,"E")).setValues(updown); 
// };

function NSLookup(name) {
  var api_url = 'https://dns.google.com/resolve'; // Google Pubic DNS API Url
  var type = 'A'; // Type of record to fetch, A, AAAA, MX, CNAME, TXT, ANY
  var requestUrl = api_url + '?name=' + name + '&type=' + type; // Build request URL
  var response = UrlFetchApp.fetch(requestUrl); // Fetch the reponse from API
  var responseText = response.getContentText(); // Get the response text
  var json = JSON.parse(responseText); // Parse the JSON text  
  
  var answers = json.Answer.map(function(ans) {
    return ans.data
  }).join('\n'); // Get the values
  return answers;
}

// Another NSLookup example
//function NSLookup1(dn,type, first) {
////  dn="mis.twse.com.tw";
//  var url = "https://cloudflare-dns.com/dns-query?ct=application/dns-json&name=" + dn + "&type=" + type;
//  var result = UrlFetchApp.fetch(url,{muteHttpExceptions:true});
//  var rc = result.getResponseCode();
//  var response = result.getContentText();
//  if (rc !== 200) {
//    throw new Error( response.message );
//  }
//  var resultData = JSON.parse(response);
//  var result = [];
//  return resultData.Answer[0].data;
//  for(var i = 0; i < resultData.Answer.length; i++) {
//    result.push(resultData.Answer[i].data);
//  }
//  if(first) {
//    
//    return result.length > 0 ? result[0] : "";
//  }
//  return result.join(" ");
//}



