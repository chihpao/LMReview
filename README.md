# LMReview

這是給內部同仁使用的檔案審查小工具，目標是快速標記檔案、生成提示詞、輸出 Word 報告。

## 使用流程
1) 點「📂」開啟 `input` 資料夾，把檔案放進去
2) 左側替未標記檔案加上【標準/範本/待審】標籤
3) 中間選擇待審檔案，生成提示詞並貼到 NotebookLM
4) 把 AI 回覆貼回來，按「輸出 Word 報告」
5) 或直接用「從剪貼簿輸出 / 自動監聽剪貼簿」

## 執行（原始碼）
1) 安裝相依套件  
   `python -m pip install -r requirements.txt`
2) 執行  
   `python notebooklm_single_folder_flow.py`

## 打包成 EXE（同仁無需安裝 Python）
1) 執行 `build_exe.ps1`（PowerShell）或 `build_exe.bat`
2) 產出位置：`dist\LMReview.exe`
3) 請把 exe 放在可寫入的資料夾，程式會在同層建立專案資料夾與 `logs`  
   若所在資料夾為唯讀，程式會改用 `使用者資料夾\LMReview_Review`
