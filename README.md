# twStockinSpreadsheet
在 Google Spreadsheet 中擷取台灣股市上市/上櫃股票價格

擷取台灣證券交易所之股價資料
參考 Asoul/tsrtc 在台灣股票即時爬蟲 Taiwan Stock Exchange Real Time Crawler 撰寫的 getStockInfo.jsp 的分析資料

# 參考資料

* [用Google試算表追蹤百檔以上台股(上市/上櫃/興櫃)股價](https://wiki0918.pixnet.net/blog/post/222332253-用google試算表取得台股%28上市-上櫃%29股票報價)
* [Asoul/tsrtc 的證交所 API Document （偽）](https://github.com/Asoul/tsrtc)
* [Google 參考文件](https://developers.google.com/apps-script/reference/spreadsheet)

# 使用方法

* 直接開啟Google Spreadsheet [股票觀察_熱鍵分享板](https://docs.google.com/spreadsheets/d/1K0OgjeL3uMZZ7JvbjE-MbkbKldgjjPEt9DYyzuuuelU/edit?usp=sharing)的分享檔案
* 按 Ctrl + Alt+ Shift + 1 即可執行執行
* 若無法從台灣證券交易所獲得資料，在 M1 儲存格會留下查詢的字串，M2 會顯示重新執行的熱鍵
* 按 Ctrl + Alt + Shift + 0 再次向台灣證券交易所，索取資料
(請注意，TWSE 有 request limit, 每 5 秒鐘 3 個 request，超過的話會被 ban 掉，請自行注意)

## 建立個人的觀察檔案

* 直接開啟Google Spreadsheet [股票觀察_熱鍵分享板](https://docs.google.com/spreadsheets/d/1K0OgjeL3uMZZ7JvbjE-MbkbKldgjjPEt9DYyzuuuelU/edit?usp=sharing)的分享檔案
* 點選檔案=>建立複本，即可以建立個人專屬的

