# tmdb-importer
將 Netflix, Disney Plus, Amazon Prime Video, Apple TV+ 等香港版的影音資料如標題簡介上載到 TMDB. 

全部由 ChatGPT 編寫，是自家用


運行前需要安裝 python

之後執行

pip install pandas selenium webdriver_manager beautifulsoup4 openpyxl psutil

安裝完後輸入

python ./tmdb-importer

就可以使用

之後可以輸入影音網站網址，或者輸入 excel 將 excel 資料上載

而 TMDB 可以輸入 TMDB 網址，或直接按 ENTER，這樣的話就會將影音網站的資料儲存在 excel

用家可以將 excel 內容修改後再上載到 TMDB

注意:如果有 SSL error 請忽略，沒有問題的


如果要功能全面的版本請使用 https://github.com/fzlins/TMDB-Import

支援影音網站更多, 功能更多更全面
