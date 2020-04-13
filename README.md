# 在 Google Spreadsheet 中擷取台灣股市上市/上櫃股票價格

在  Google SPreadsheet，透過巨集熱鍵，擷取台灣證券交易所之股價資料

參考 Asoul/tsrtc 在台灣股票即時爬蟲 Taiwan Stock Exchange Real Time Crawler 撰寫的 getStockInfo.jsp 的分析資料

## 參考資料

* [用Google試算表追蹤百檔以上台股(上市/上櫃/興櫃)股價](https://wiki0918.pixnet.net/blog/post/222332253-用google試算表取得台股%28上市-上櫃%29股票報價)
* [Asoul/tsrtc 的證交所 API Document （偽）](https://github.com/Asoul/tsrtc)
* [Google 參考文件](https://developers.google.com/apps-script/reference/spreadsheet)

## 影片教學

* [Google試算表整理台股投資 | 透過 Javascript 程式，讓 google 試算表自動更新股票價格，掌握個人投資現況](https://youtu.be/VlPjjLtXzY0)

## 使用方法

* 直接開啟Google Spreadsheet [股票觀察_熱鍵分享板](https://docs.google.com/spreadsheets/d/1K0OgjeL3uMZZ7JvbjE-MbkbKldgjjPEt9DYyzuuuelU/edit?usp=sharing)的分享檔案
* 需要登入 Google 帳號，並授予 stockPriceAll 這一個應用程序，存取 google 帳號的部分權限
* 點擊「執行」圖示或按 Ctrl + Alt+ Shift + 1 即可執行執行
* 若無法從台灣證券交易所獲得資料，在 A3 會出現重試字樣，並在在 M1 儲存格會留下查詢的字串，M2 會顯示重新執行的熱鍵
* 點擊「重試」圖示或按 Ctrl + Alt + Shift + 0 再次向台灣證券交易所，索取資料
(請注意，TWSE 有 request limit, 每 5 秒鐘 3 個 request，超過的話會被 ban 掉，請自行注意)

## 建立個人的觀察檔案

* 直接開啟Google Spreadsheet [股票觀察_熱鍵分享板](https://docs.google.com/spreadsheets/d/1K0OgjeL3uMZZ7JvbjE-MbkbKldgjjPEt9DYyzuuuelU/edit?usp=sharing)的分享檔案
* 點選檔案=>建立複本，即可以建立個人專屬的檔案

## 異常處理

檢查 Google Spreadsheet 的巨集設定，確定 setStocksPrice, retry_setStocksPrice 被正確載入，並完成了熱鍵的設定

* setStocksPrice 的熱鍵為 Ctrl+Alt+Shift+1
* retry_setStocksPrice 的熱鍵為 Ctrl+Alt+Shift+0

若函式未載入，請透過 工具=>巨集=>匯入 ，將上述兩個函數匯入。並透過 工具=>巨集=>管理巨集 完成熱鍵的設定。

## 函數說明

共 4 個函數，分別為 setStocksPrice()、retry_setStocksPrice()、getStockQueryString() 、getStockPrice(stocksID)

* setStocksPrice() 主要的程式，呼叫 getStockQueryString() 得到查詢字串，呼叫 getStockPrice(stocksID) 取得包含股票資訊的 json 檔，並完成價格的更新
* retry_setStocksPrice() 重試的程式，呼叫 getStockPrice(stocksID) 取得包含股票資訊的 json 檔並完成價格的更新
* getStockQueryString() 用來組成「股票查詢字串」的函數，由表格的 A4 欄位往下取得，D欄 狀態設定為 1 的資料。 A4 往下至第一個空白的 A 欄位即停止
* getStockPrice(stocksID) 向查詢台灣證交所的 getStockInfo.jsp 查詢，並得到 json 資料回傳





