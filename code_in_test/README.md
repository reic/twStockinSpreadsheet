# 測試中的 macro js

優化 Google SPreadsheet，擷取台灣證券交易所之股價資料

## 函數說明

共 4 個函數，分別為 FastMode_construct()、re_Construct()、 clearStocksPrice() 、preSetQuery()、getStockPrice()、setNewStockPrice()

* FastMode_construct、re_Construct 為主要的運作程式，用來控制取得資料使用。
* preSetQuery 完成股票資料的初始設定，回傳值為陣列  **[price, querystring]** ，price 會留下獲利了結的的股票價格, querystring 為去除重複後的股票查詢代碼
* clearStocksPrice 為清除股票現價的函數，透過這一個函數，可以實現慢慢的清除資料的動作
* getStockPrice(url) 向查詢台灣證交所的 getStockInfo.jsp 查詢，並得到 json 資料回傳
* setNewStockPrice 更新新的股票價格至 Google spreadsheet

## 更新說明

* 20200413 在 FastMode_construct、re_Construct 導入 Google Utitlies sleep ，加入 5 秒自動重試
