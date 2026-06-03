# 🏭 產地自動分配與生管達交檢核系統 (Production Allocation & PC Verification System)

這是一個專為製造業 **生產計劃（Production Planning）** 與 **資材資深幕僚** 設計的智慧化營運工具。
成功將傳統人工比對、查閱多張對照表並手動剪貼的繁瑣作業，優化為**自動化多級匹配與交叉防錯管線**。

## 💡 核心管理價值與技術亮點

* **📊 四級優先權漏斗匹配演算法**：
  建構深度的數據檢索邏輯，系統會自動依據 `料號` ➔ `客戶簡稱` ➔ `內外單` ➔ `業務地區` 進行四階層優先級（Priority）比對，並動態計算 `起始日期` 時間戳記，自動判定每週 Forecast 訂單的初步最佳投產地。
  
* **🚨 銷售組織交叉防錯機制 (Cross-Validation)**：
  實務生管常面臨人工貼錯排程的風險。本系統在「生管達交比對」模組中導入自動稽核邏輯：當輸入之銷售組織為海外廠（如 `CN10` / `TH10`），但匹配結果為本島（`高雄`）或 `空白` 時，系統將**瞬間觸發紅色 Error 警報並列出異常明細**，達成零疏漏的生產線與資財控管。

* **✏️ 動態對照表線上維護 (Live Data Editor)**：
  整合 Streamlit Data Editor 技術，無須修改後端資料庫，生管主管即可直接在網頁上對四張核心 Mapping 表進行增刪查改（CRUD），並實現一鍵同步儲存與快取清除（Cache Clear）。

## 🛠 開發套件與架構
* **Backend**: Python (Pandas 資料清洗、DateTime 時間序列過濾、OpenPyXL 引擎)
* **Frontend**: Streamlit Web Framework (Session State 狀態保持)
