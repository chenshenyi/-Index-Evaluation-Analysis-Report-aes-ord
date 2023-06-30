# 測試檔說明

## 資料夾結構

測試檔案分成兩大類，一部分是會出現在正式專案中的資料，包括：

- 0.原始資料
- 1.各院彙整資料
- 2.各院報告書模板
- 3.各院報告書
- A 主程式.xlsm
- B 參數.xlsx

另一部份是測試專用的資料，根據測試的函數命名，包括：

- input: 預設輸入資料，受 `git` 管理，所有測試操作都不能更改此資料夾內容
- output: 輸出資料，不受 `git` 管理，修改檔案的函數測試，皆應從 `input` 資料夾讀取或複製，並將結果寫入 `output` 資料夾

## 匯入模組(module)

所有程式碼都寫在 `A. 主程式.xlsm` 中，因為 `VBA` 是一種古老的語言，所以匯入模組的方式也很古老，此處使用由 [henriquebcustodio](https://github.com/henriquebcustodio) 開發的插件。

在 `A. 主程式.xlsm` 的 `表單(User Form)` 中找到 `gitExportImport`，點擊後直接執行(F5)，詳細說明請參考 [VBA-Import-Export](https://github.com/henriquebcustodio/git-export-import-vba)。

## 發布正式版

將資料夾中 `A. 主程式.xlsm` 分別複製到 `指標報告書自動化範例檔案` 與 `新專案` 中即可。
