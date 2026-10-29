# 生物統計實習：無母數統計分析報告

本專案收錄生物統計實習之三組資料集無母數統計檢定分析（Mann-Whitney U 檢定、Kruskal-Wallis H 檢定及事後檢定 Dunn's test）、繪圖與自動化報表生成程式碼及最終報告。

---

## 📁 檔案結構說明

| 檔案名稱 | 說明 |
| :--- | :--- |
| `generate_reports.py` | 自動執行統計檢定、繪製圖表並匯出圖表與摘要的 Python 腳本 |
| `statistical_summary_results.csv` | 三份資料集統計檢定結果彙整表（含統計量、p 值、效應值與結論） |
| `生統實習_三份資料統計分析報告.docx` | 完整中文分析報告（Word 檔） |
| `生統實習_三份資料統計分析報告.pdf` | 完整中文分析報告（PDF 檔） |
| `Biostatistics_Statistical_Analysis_Report.docx` | 英文版統計分析報告（Word 檔） |
| `Mann whitney.xlsx` | 資料集 1：乳癌 (BC) 與肺腺癌 (LUAD) 病患之 HER2 表現量 |
| `KW_Blood type.xlsx` | 資料集 2：不同血型族群 (A, B, O, AB) 之身高資料 |
| `KW_Pro_salary.xlsx` | 資料集 3：四大職業運動聯盟 (NBA, MLB, NFL, WNBA) 球員薪資與年齡 |
| `fig1_mann_whitney_her2.png` | 圖 1：HER2 表現量在乳癌與肺腺癌之分布比較圖 |
| `fig2_kw_blood_type_height.png` | 圖 2：各血型族群身高之分布比較圖 |
| `fig3_kw_pro_salary.png` | 圖 3：四大職業運動聯盟球員薪資分布比較圖 |
| `fig4_kw_pro_age.png` | 圖 4：四大職業運動聯盟球員年齡分布比較圖 |

---

## 📊 分析結果摘要

| 資料集 | 分析變數 | 檢定方法 | 統計量與 p 值 | 結論 |
| :--- | :--- | :--- | :--- | :--- |
| **Mann whitney.xlsx** | HER2 表現量 | Mann-Whitney U | U=1.0, *p* = 0.0009 (***) | 乳癌患者 HER2 表現量顯著高於肺腺癌 (*r* = 0.806) |
| **KW_Blood type.xlsx** | 身高 (cm) | Kruskal-Wallis H | H(3)=1.507, *p* = 0.6806 (ns) | 各血型族群間身高無顯著差異 |
| **KW_Pro_salary.xlsx** | 年薪 (萬美元) | Kruskal-Wallis H & Dunn | H(3)=27.611, *p* < 0.0001 (***) | WNBA 薪資顯著低於 NBA, MLB, NFL (*p* < 0.05) |
| **KW_Pro_salary.xlsx** | 年齡 (歲) | Kruskal-Wallis H & Dunn | H(3)=8.831, *p* = 0.0316 (*) | MLB 球員年齡顯著高於 NFL 球員 (*p* = 0.021) |

---

## 🛠️ 環境與執行方式

### 需求套件
```bash
pip install pandas numpy scipy matplotlib seaborn openpyxl python-docx
```

### 執行分析
```bash
python generate_reports.py
```
執行後將自動生成 `fig1` ~ `fig4` 視覺化圖表與 `statistical_summary_results.csv`。
