# -*- coding: utf-8 -*-
"""
生統實習資料統計分析腳本 (升級視覺化版)
包含三份資料集之描述性統計、常態性與變異數同質性檢定、
非母數檢定 (Mann-Whitney U、Kruskal-Wallis H)、事後比較 (Dunn's test)、
Seaborn 高解析度圖表繪製，以及生成完整 Word (.docx) 分析報告。
"""

import os
import sys
import shutil
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Matplotlib 中文字型設定
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Microsoft YaHei', 'SimHei', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False
sns.set_theme(style="whitegrid", font="Microsoft JhengHei")

base_dir = r"c:\Users\蘇子恩\OneDrive\文件\生統實習\20261015"
artifact_dir = r"C:\Users\蘇子恩\.gemini\antigravity-ide\brain\0852ec8f-048d-466b-8f40-d505f4418f99"
os.makedirs(artifact_dir, exist_ok=True)
os.chdir(base_dir)

# ==========================================
# 1. 讀取並分析 DATASET 1: Mann whitney.xlsx
# ==========================================
print("--> 分析 Dataset 1: Mann whitney.xlsx")
df_mw = pd.read_excel("Mann whitney.xlsx", sheet_name="工作表1")
bc_vals = df_mw[df_mw['annotation'] == 'BC']['HER2_level'].values
luad_vals = df_mw[df_mw['annotation'] == 'LUAD']['HER2_level'].values

desc_mw = df_mw.groupby('annotation')['HER2_level'].agg(
    n='count',
    mean='mean',
    std='std',
    median='median',
    q25=lambda x: np.percentile(x, 25),
    q75=lambda x: np.percentile(x, 75),
    min='min',
    max='max'
).loc[['BC', 'LUAD']]
desc_mw['iqr'] = desc_mw['q75'] - desc_mw['q25']

shapiro_bc = stats.shapiro(bc_vals)
shapiro_luad = stats.shapiro(luad_vals)
levene_mw = stats.levene(bc_vals, luad_vals)

mw_res = stats.mannwhitneyu(bc_vals, luad_vals, alternative='two-sided')
df_mw['rank'] = stats.rankdata(df_mw['HER2_level'])
rank_sum_bc = df_mw[df_mw['annotation'] == 'BC']['rank'].sum()
rank_sum_luad = df_mw[df_mw['annotation'] == 'LUAD']['rank'].sum()
mean_rank_bc = df_mw[df_mw['annotation'] == 'BC']['rank'].mean()
mean_rank_luad = df_mw[df_mw['annotation'] == 'LUAD']['rank'].mean()

n1, n2 = len(bc_vals), len(luad_vals)
U_bc = rank_sum_bc - n1 * (n1 + 1) / 2 # 71
U_luad = rank_sum_luad - n2 * (n2 + 1) / 2 # 1
U_min = min(U_bc, U_luad)
r_biserial = 1 - (2 * U_luad) / (n1 * n2)

mu_u = n1 * n2 / 2
sigma_u = np.sqrt(n1 * n2 * (n1 + n2 + 1) / 12)
z_val = (U_bc - mu_u) / sigma_u
effect_r = z_val / np.sqrt(n1 + n2)
ttest_res = stats.ttest_ind(bc_vals, luad_vals, equal_var=False)

# 繪製圖 1 (調整間距避免文字遮蔽)
fig, ax = plt.subplots(figsize=(7.5, 6.5), dpi=300)
palette = {'BC': '#D9534F', 'LUAD': '#428BCA'}
sns.boxplot(x='annotation', y='HER2_level', data=df_mw, hue='annotation', palette=palette, width=0.45,
            boxprops=dict(alpha=0.75), showmeans=True, legend=False,
            meanprops={"marker":"o", "markerfacecolor":"white", "markeredgecolor":"black", "markersize":"8"},
            ax=ax)
sns.stripplot(x='annotation', y='HER2_level', data=df_mw, hue='annotation', palette=['#333333', '#333333'],
              alpha=0.75, size=8, jitter=0.15, legend=False, ax=ax)

# 顯著性標註放在上方
y_max = df_mw['HER2_level'].max()
y_bracket = 710
ax.plot([0, 0, 1, 1], [y_bracket, y_bracket + 20, y_bracket + 20, y_bracket], lw=1.5, c='black')
ax.text(0.5, y_bracket + 28, f"*** 極顯著差異 (p = {mw_res.pvalue:.4f})\nMann-Whitney U = {U_min:.1f} (或 {U_bc:.1f})", 
        ha='center', va='bottom', fontsize=11, fontweight='bold', color='#B22222')

ax.set_title("乳癌 (BC) 與肺腺癌 (LUAD) 患者 HER2 表現量比較\n(Mann-Whitney U Test)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("癌症類別 (Annotation)", fontsize=12, labelpad=10)
ax.set_ylabel("HER2 表現量 (HER2 Level)", fontsize=12, labelpad=10)
ax.set_xticks([0, 1])
ax.set_xticklabels(["乳癌 (BC, n=8)", "肺腺癌 (LUAD, n=9)"], fontsize=11)
ax.set_ylim(-20, 880)

# 說明框置於右側 LUAD 上方空曠處
textstr = f"BC (n=8):\n  中位數 = {desc_mw.loc['BC','median']:.1f}, IQR = {desc_mw.loc['BC','iqr']:.1f}\n  平均秩次 = {mean_rank_bc:.2f}\n" \
          f"LUAD (n=9):\n  中位數 = {desc_mw.loc['LUAD','median']:.1f}, IQR = {desc_mw.loc['LUAD','iqr']:.1f}\n  平均秩次 = {mean_rank_luad:.2f}\n" \
          f"檢定結果:\n  Mann-Whitney U = {U_min:.1f} (或 {U_bc:.1f})\n  p = {mw_res.pvalue:.4f} < 0.001\n" \
          f"效應值 (Rank-biserial r) = {r_biserial:.3f}"
props = dict(boxstyle='round,pad=0.5', facecolor='#F7F9FA', edgecolor='#BDC3C7', alpha=0.92)
ax.text(0.55, 0.75, textstr, transform=ax.transAxes, fontsize=9.5, verticalalignment='top', bbox=props)

plt.tight_layout()
fig1_path = os.path.join(base_dir, "fig1_mann_whitney_her2.png")
fig.savefig(fig1_path, dpi=300)
shutil.copy(fig1_path, os.path.join(artifact_dir, "fig1_mann_whitney_her2.png"))
plt.close(fig)
print("已生成圖 1:", fig1_path)


# ==========================================
# 2. 讀取並分析 DATASET 2: KW_Blood type.xlsx
# ==========================================
print("--> 分析 Dataset 2: KW_Blood type.xlsx")
df_bt = pd.read_excel("KW_Blood type.xlsx", sheet_name="工作表1")
order_bt = ['A', 'B', 'O', 'AB']

desc_bt = df_bt.groupby('blood type')['Height (cm)'].agg(
    n='count',
    mean='mean',
    std='std',
    median='median',
    q25=lambda x: np.percentile(x, 25),
    q75=lambda x: np.percentile(x, 75),
    min='min',
    max='max'
).loc[order_bt]
desc_bt['iqr'] = desc_bt['q75'] - desc_bt['q25']

shapiro_bt = {bt: stats.shapiro(df_bt[df_bt['blood type'] == bt]['Height (cm)']) for bt in order_bt}
groups_bt = [df_bt[df_bt['blood type'] == bt]['Height (cm)'].values for bt in order_bt]
levene_bt = stats.levene(*groups_bt)

df_bt['rank'] = stats.rankdata(df_bt['Height (cm)'])
mean_ranks_bt = df_bt.groupby('blood type')['rank'].mean().loc[order_bt]
kw_bt = stats.kruskal(*groups_bt)

N_bt = len(df_bt)
k_bt = len(order_bt)
epsilon2_bt = max(0.0, (kw_bt.statistic - k_bt + 1) / (N_bt - k_bt))
anova_bt = stats.f_oneway(*groups_bt)

fig, ax = plt.subplots(figsize=(7.5, 6), dpi=300)
palette_bt = {'A': '#5DADE2', 'B': '#58D68D', 'O': '#F5B041', 'AB': '#AF7AC5'}
sns.boxplot(x='blood type', y='Height (cm)', data=df_bt, order=order_bt, hue='blood type', palette=palette_bt, width=0.45,
            boxprops=dict(alpha=0.75), showmeans=True, legend=False,
            meanprops={"marker":"o", "markerfacecolor":"white", "markeredgecolor":"black", "markersize":"8"},
            ax=ax)
sns.stripplot(x='blood type', y='Height (cm)', data=df_bt, order=order_bt, color='black', alpha=0.65, size=7, jitter=0.15, ax=ax)

ax.text(0.5, 0.94, f"ns (無統計顯著差異, p = {kw_bt.pvalue:.3f})\nKruskal-Wallis H(3) = {kw_bt.statistic:.3f}", 
        ha='center', va='top', transform=ax.transAxes, fontsize=11, fontweight='bold', color='#444444')

ax.set_title("不同血型族群之身高分佈比較\n(Kruskal-Wallis Test)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("血型 (Blood Type)", fontsize=12, labelpad=10)
ax.set_ylabel("身高 (Height, cm)", fontsize=12, labelpad=10)
ax.set_xticks(range(4))
ax.set_xticklabels([f"A 型 (n=12)", f"B 型 (n=10)", f"O 型 (n=9)", f"AB 型 (n=13)"], fontsize=11)
ax.set_ylim(145, 198)

textstr_bt = f"各組平均秩次 (Mean Rank): A={mean_ranks_bt['A']:.2f}, B={mean_ranks_bt['B']:.2f}, O={mean_ranks_bt['O']:.2f}, AB={mean_ranks_bt['AB']:.2f}\n" \
             f"Kruskal-Wallis 檢定: H = {kw_bt.statistic:.3f}, df = {k_bt-1}, p = {kw_bt.pvalue:.3f} > 0.05\n" \
             f"單因子 ANOVA 對照: F = {anova_bt.statistic:.3f}, p = {anova_bt.pvalue:.3f} > 0.05\n" \
             f"結論: 四種血型之間的身高無統計顯著差異 (血型與身高無關聯)"
props = dict(boxstyle='round,pad=0.5', facecolor='#F7F9FA', edgecolor='#BDC3C7', alpha=0.92)
ax.text(0.03, 0.14, textstr_bt, transform=ax.transAxes, fontsize=9.5, verticalalignment='top', bbox=props)

plt.tight_layout()
fig2_path = os.path.join(base_dir, "fig2_kw_blood_type_height.png")
fig.savefig(fig2_path, dpi=300)
shutil.copy(fig2_path, os.path.join(artifact_dir, "fig2_kw_blood_type_height.png"))
plt.close(fig)
print("已生成圖 2:", fig2_path)


# ==========================================
# 3. 讀取並分析 DATASET 3: KW_Pro_salary.xlsx
# ==========================================
print("--> 分析 Dataset 3: KW_Pro_salary.xlsx")
df_pro = pd.read_excel("KW_Pro_salary.xlsx", sheet_name="工作表1")
order_pro = ['NBA', 'MLB', 'NFL', 'WNBA']

df_pro['Salary_Million'] = df_pro['Salary (Annual)'] / 1e6
desc_pro = df_pro.groupby('profession')['Salary (Annual)'].agg(
    n='count',
    mean='mean',
    std='std',
    median='median',
    q25=lambda x: np.percentile(x, 25),
    q75=lambda x: np.percentile(x, 75),
    min='min',
    max='max'
).loc[order_pro]
desc_pro['iqr'] = desc_pro['q75'] - desc_pro['q25']

shapiro_pro = {sport: stats.shapiro(df_pro[df_pro['profession'] == sport]['Salary (Annual)']) for sport in order_pro}
groups_pro_sal = [df_pro[df_pro['profession'] == sport]['Salary (Annual)'].values for sport in order_pro]
levene_pro = stats.levene(*groups_pro_sal)

df_pro['rank_sal'] = stats.rankdata(df_pro['Salary (Annual)'])
mean_ranks_pro = df_pro.groupby('profession')['rank_sal'].mean().loc[order_pro]
kw_pro = stats.kruskal(*groups_pro_sal)

N_pro = len(df_pro)
k_pro = len(order_pro)
epsilon2_pro = (kw_pro.statistic - k_pro + 1) / (N_pro - k_pro)

dunn_sal_results = []
n_pro_each = 10
for i in range(len(order_pro)):
    for j in range(i+1, len(order_pro)):
        s1, s2 = order_pro[i], order_pro[j]
        r1, r2 = mean_ranks_pro[s1], mean_ranks_pro[s2]
        se = np.sqrt((N_pro * (N_pro + 1) / 12.0) * (1/n_pro_each + 1/n_pro_each))
        z = (r1 - r2) / se
        p_raw = 2 * (1 - stats.norm.cdf(abs(z)))
        p_bonf = min(1.0, p_raw * 6)
        dunn_sal_results.append({
            'Comparison': f"{s1} vs {s2}",
            'Rank_Diff': r1 - r2,
            'Z': z,
            'p_raw': p_raw,
            'p_bonf': p_bonf,
            'Significant': p_bonf < 0.05
        })
df_dunn_sal = pd.DataFrame(dunn_sal_results)

desc_age = df_pro.groupby('profession')['Age'].agg(
    n='count',
    mean='mean',
    std='std',
    median='median',
    q25=lambda x: np.percentile(x, 25),
    q75=lambda x: np.percentile(x, 75),
    min='min',
    max='max'
).loc[order_pro]
desc_age['iqr'] = desc_age['q75'] - desc_age['q25']

groups_pro_age = [df_pro[df_pro['profession'] == sport]['Age'].values for sport in order_pro]
kw_age = stats.kruskal(*groups_pro_age)
df_pro['rank_age'] = stats.rankdata(df_pro['Age'])
mean_ranks_age = df_pro.groupby('profession')['rank_age'].mean().loc[order_pro]

dunn_age_results = []
for i in range(len(order_pro)):
    for j in range(i+1, len(order_pro)):
        s1, s2 = order_pro[i], order_pro[j]
        r1, r2 = mean_ranks_age[s1], mean_ranks_age[s2]
        se = np.sqrt((N_pro * (N_pro + 1) / 12.0) * (1/n_pro_each + 1/n_pro_each))
        z = (r1 - r2) / se
        p_raw = 2 * (1 - stats.norm.cdf(abs(z)))
        p_bonf = min(1.0, p_raw * 6)
        dunn_age_results.append({
            'Comparison': f"{s1} vs {s2}",
            'Rank_Diff': r1 - r2,
            'Z': z,
            'p_raw': p_raw,
            'p_bonf': p_bonf,
            'Significant': p_bonf < 0.05
        })
df_dunn_age = pd.DataFrame(dunn_age_results)

male_sal = df_pro[df_pro['Gender'] == 1]['Salary (Annual)'].values
female_sal = df_pro[df_pro['Gender'] == 2]['Salary (Annual)'].values
mw_gender = stats.mannwhitneyu(male_sal, female_sal, alternative='two-sided')

# 繪製圖 3: 薪資比較 (調整佈局與 Y 軸上限避免文字與點重疊)
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6.5), dpi=300)
palette_pro = {'NBA': '#E67E22', 'MLB': '#2980B9', 'NFL': '#27AE60', 'WNBA': '#8E44AD'}

# 子圖 1: 百萬美元線性尺度
sns.boxplot(x='profession', y='Salary_Million', data=df_pro, order=order_pro, hue='profession', palette=palette_pro, width=0.45,
            boxprops=dict(alpha=0.75), showmeans=True, legend=False,
            meanprops={"marker":"o", "markerfacecolor":"white", "markeredgecolor":"black", "markersize":"7"},
            ax=ax1)
sns.stripplot(x='profession', y='Salary_Million', data=df_pro, order=order_pro, color='black', alpha=0.65, size=6, jitter=0.15, ax=ax1)
ax1.set_title("四大職業聯盟頂尖球員年薪比較 (線性百萬美元尺度)", fontsize=13, fontweight='bold', pad=12)
ax1.set_xlabel("職業運動聯盟", fontsize=11, labelpad=8)
ax1.set_ylabel("年薪 (百萬美元, Million USD)", fontsize=11, labelpad=8)
ax1.set_xticks(range(4))
ax1.set_xticklabels(order_pro, fontsize=11)
ax1.set_ylim(-3, 85)
ax1.text(0.5, 0.94, f"Kruskal-Wallis H(3) = {kw_pro.statistic:.2f}\np = {kw_pro.pvalue:.2e} (*** 極顯著差異)",
         ha='center', va='top', transform=ax1.transAxes, fontsize=11, fontweight='bold', color='#B22222')

# 子圖 2: 對數尺度 (Log10 Scale)
sns.boxplot(x='profession', y='Salary (Annual)', data=df_pro, order=order_pro, hue='profession', palette=palette_pro, width=0.45,
            boxprops=dict(alpha=0.75), showmeans=True, legend=False,
            meanprops={"marker":"o", "markerfacecolor":"white", "markeredgecolor":"black", "markersize":"7"},
            ax=ax2)
sns.stripplot(x='profession', y='Salary (Annual)', data=df_pro, order=order_pro, color='black', alpha=0.65, size=6, jitter=0.15, ax=ax2)
ax2.set_yscale('log')
ax2.set_title("四大職業聯盟頂尖球員年薪比較 (對數尺度 Log10)", fontsize=13, fontweight='bold', pad=12)
ax2.set_xlabel("職業運動聯盟", fontsize=11, labelpad=8)
ax2.set_ylabel("年薪 (美元, 對數座標 Log10 Scale)", fontsize=11, labelpad=8)
ax2.set_xticks(range(4))
ax2.set_xticklabels(order_pro, fontsize=11)

y_ref = 1.3e8
ax2.plot([0, 0, 3, 3], [y_ref * 0.7, y_ref, y_ref, y_ref * 0.7], lw=1.3, c='black')
ax2.text(1.5, y_ref * 1.15, "Dunn's post-hoc: WNBA 與三大聯盟均有極顯著落差 (p < 0.05)", ha='center', fontsize=10, fontweight='bold', color='#B22222')
ax2.set_ylim(4e4, 3.5e8)

plt.tight_layout()
fig3_path = os.path.join(base_dir, "fig3_kw_pro_salary.png")
fig.savefig(fig3_path, dpi=300)
shutil.copy(fig3_path, os.path.join(artifact_dir, "fig3_kw_pro_salary.png"))
plt.close(fig)
print("已生成圖 3:", fig3_path)


# 繪製圖 4: 年齡比較
fig, ax = plt.subplots(figsize=(7.5, 6.2), dpi=300)
sns.boxplot(x='profession', y='Age', data=df_pro, order=order_pro, hue='profession', palette=palette_pro, width=0.45,
            boxprops=dict(alpha=0.75), showmeans=True, legend=False,
            meanprops={"marker":"o", "markerfacecolor":"white", "markeredgecolor":"black", "markersize":"8"},
            ax=ax)
sns.stripplot(x='profession', y='Age', data=df_pro, order=order_pro, color='black', alpha=0.65, size=7, jitter=0.15, ax=ax)

# 標註年齡事後比較 (MLB vs NFL)
y_age_max = 42
ax.plot([1, 1, 2, 2], [y_age_max + 1.5, y_age_max + 2.5, y_age_max + 2.5, y_age_max + 1.5], lw=1.3, c='black')
ax.text(1.5, y_age_max + 2.8, "* MLB vs NFL (Bonferroni p = 0.021)", ha='center', va='bottom', fontsize=10, fontweight='bold', color='#B22222')

ax.set_title("四大職業聯盟頂尖球員年齡分佈比較\n(Kruskal-Wallis Test)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("職業運動聯盟", fontsize=12, labelpad=10)
ax.set_ylabel("年齡 (Age, 歲)", fontsize=12, labelpad=10)
ax.set_xticks(range(4))
ax.set_xticklabels(order_pro, fontsize=11)
ax.set_ylim(20, 50)

textstr_age = f"Kruskal-Wallis 檢定: H = {kw_age.statistic:.3f}, df = 3, p = {kw_age.pvalue:.4f} < 0.05 (* 達顯著差異)\n" \
              f"各組平均年齡: MLB (33.6歲) > NBA (30.2歲) > WNBA (29.8歲) > NFL (27.6歲)\n" \
              f"Dunn 事後檢定: MLB 球員顯著年長於 NFL 球員 (Bonferroni p = 0.021)\n" \
              f"其餘聯盟兩兩之間無統計學顯著年齡差異。"
props = dict(boxstyle='round,pad=0.5', facecolor='#F7F9FA', edgecolor='#BDC3C7', alpha=0.92)
ax.text(0.03, 0.16, textstr_age, transform=ax.transAxes, fontsize=9.5, verticalalignment='top', bbox=props)

plt.tight_layout()
fig4_path = os.path.join(base_dir, "fig4_kw_pro_age.png")
fig.savefig(fig4_path, dpi=300)
shutil.copy(fig4_path, os.path.join(artifact_dir, "fig4_kw_pro_age.png"))
plt.close(fig)
print("已生成圖 4:", fig4_path)


# ==========================================
# 4. 生成 Word (.docx) 正式統計分析交接報告
# ==========================================
print("--> 正在建置 Word (.docx) 正式報告...")
doc = docx.Document()

for section in doc.sections:
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

def set_run_font(run, font_name="微軟正黑體", size_pt=11, bold=False, color_rgb=(0, 0, 0)):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)

def style_table(table, col_widths=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_tr = table.rows[0]._tr.get_or_add_trPr()
    header_tr.append(OxmlElement('w:tblHeader'))
    
    for row_idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), '360')
        trPr.append(trHeight)
        
        for col_idx, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            
            if col_widths and col_idx < len(col_widths):
                cell.width = col_widths[col_idx]
                
            shd = OxmlElement('w:shd')
            if row_idx == 0:
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), '1F4E79')
                tcPr.append(shd)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        set_run_font(r, "微軟正黑體", 10, bold=True, color_rgb=(255, 255, 255))
            else:
                if row_idx % 2 == 1:
                    shd.set(qn('w:fill'), 'F2F4F8')
                else:
                    shd.set(qn('w:fill'), 'FFFFFF')
                tcPr.append(shd)
                for p in cell.paragraphs:
                    if col_idx == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        set_run_font(r, "微軟正黑體", 9.5, bold=False, color_rgb=(51, 51, 51))

# 標題
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.paragraph_format.space_before = Pt(10)
title_p.paragraph_format.space_after = Pt(4)
r_title = title_p.add_run("生物統計學實習：非母數統計分析綜合報告")
set_run_font(r_title, "微軟正黑體", 22, bold=True, color_rgb=(31, 78, 121))

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_p.paragraph_format.space_after = Pt(18)
r_sub = sub_p.add_run("Mann-Whitney U 檢定與 Kruskal-Wallis H 檢定之應用與學術解讀")
set_run_font(r_sub, "微軟正黑體", 13, bold=False, color_rgb=(89, 89, 89))

meta_p = doc.add_paragraph()
meta_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
meta_p.paragraph_format.space_after = Pt(20)
r_meta = meta_p.add_run("分析日期：2026 年 10 月 15 日  |  分析檔案：Mann whitney.xlsx、KW_Blood type.xlsx、KW_Pro_salary.xlsx")
set_run_font(r_meta, "微軟正黑體", 9.5, bold=False, color_rgb=(120, 120, 120))

# 執行摘要
h1_1 = doc.add_heading(level=1)
r_h1_1 = h1_1.add_run("【執行摘要】三份資料統計分析核心結論一覽表")
set_run_font(r_h1_1, "微軟正黑體", 15, bold=True, color_rgb=(31, 78, 121))

p_sum = doc.add_paragraph()
p_sum.paragraph_format.line_spacing = 1.25
p_sum.paragraph_format.space_after = Pt(10)
r_sum = p_sum.add_run(
    "本實習分析涵蓋醫學生物標記、生理特徵與運動經濟學三項獨立主題，重點評估非獨立或非常態/小樣本連續變數之組間差異。"
    "在進行統計檢定前，均已嚴謹執行 Shapiro-Wilk 常態性檢定與 Levene 變異數同質性檢定，並據此選用最適切之統計檢定法："
)
set_run_font(r_sum, "微軟正黑體", 11)

table_sum = doc.add_table(rows=4, cols=6)
headers_sum = ["資料集檔案", "研究主題與應變數", "自變數分組", "檢定方法", "統計量與 p 值", "統計與實務結論"]
for i, h in enumerate(headers_sum):
    table_sum.cell(0, i).paragraphs[0].text = h

summary_rows = [
    ["Mann whitney.xlsx", "腫瘤 HER2 表現量", "乳癌 BC (n=8)\n肺腺癌 LUAD (n=9)", "Mann-Whitney U 檢定\n(雙尾檢定)", 
     f"U = {U_min:.1f} (或 {U_bc:.1f})\np = {mw_res.pvalue:.4f} (***)\nEffect r = {r_biserial:.3f}", 
     "極顯著差異 (p < 0.001)。乳癌患者 HER2 表現量 (中位數 446.5) 極顯著高於肺腺癌 (中位數 118.0)。"],
    ["KW_Blood type.xlsx", "成人身高 (cm)", "A型 (12), B型 (10),\nO型 (9), AB型 (13)", "Kruskal-Wallis 檢定\n(k=4 組獨立樣本)", 
     f"H(3) = {kw_bt.statistic:.3f}\np = {kw_bt.pvalue:.3f} (ns)\nε² = 0.000", 
     "無顯著差異 (p = 0.681 > 0.05)。不同血型族群在身高表現上無統計學差異，血型與身高無關聯。"],
    ["KW_Pro_salary.xlsx", "職業球員年薪\n及年齡", "NBA (10), MLB (10),\nNFL (10), WNBA (10)", "Kruskal-Wallis 檢定\n& Dunn 事後多重比較", 
     f"年薪: H(3) = {kw_pro.statistic:.2f}, p < 0.001\n年齡: H(3) = {kw_age.statistic:.2f}, p = 0.032", 
     "年薪呈極顯著聯盟差異 (p < 0.001)。Dunn 事後檢定證實 WNBA 薪資顯著低於三大男子聯盟 (p < 0.05)。年齡方面 MLB 顯著年長於 NFL (p = 0.021)。"]
]

for row_idx, r_data in enumerate(summary_rows, start=1):
    for col_idx, text in enumerate(r_data):
        table_sum.cell(row_idx, col_idx).paragraphs[0].text = text

style_table(table_sum, [Inches(1.3), Inches(1.1), Inches(1.2), Inches(1.2), Inches(1.3), Inches(1.9)])
doc.add_paragraph().paragraph_format.space_after = Pt(14)

# 第一部分
h1_2 = doc.add_heading(level=1)
r_h1_2 = h1_2.add_run("第一部分：乳癌 (BC) 與肺腺癌 (LUAD) 患者 HER2 表現量分析")
set_run_font(r_h1_2, "微軟正黑體", 15, bold=True, color_rgb=(31, 78, 121))

p = doc.add_paragraph()
r = p.add_run("1.1 研究背景與資料概況")
set_run_font(r, "微軟正黑體", 12, bold=True, color_rgb=(51, 51, 51))

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.25
r = p.add_run(
    "本資料集 (Mann whitney.xlsx) 旨在探討乳癌 (Breast Cancer, BC) 與肺腺癌 (Lung Adenocarcinoma, LUAD) 兩類腫瘤組織中，"
    "人類表皮生長因子受體 2 (HER2_level) 表現量是否存在顯著差異。總樣本數 N = 17，包含 BC 組 8 例 (group 1) 與 LUAD 組 9 例 (group 2)。"
)
set_run_font(r, "微軟正黑體", 11)

p = doc.add_paragraph()
r = p.add_run("表 1. HER2 表現量之分組描述性統計與假設檢定")
set_run_font(r, "微軟正黑體", 10.5, bold=True, color_rgb=(31, 78, 121))

table_t1 = doc.add_table(rows=3, cols=8)
headers_t1 = ["分組 (群體)", "樣本數 (n)", "平均值 ± 標準差", "中位數", "四分位距 (IQR)", "極值範圍 [Min, Max]", "常態性檢定 (W, p)", "平均秩次 (Mean Rank)"]
for i, h in enumerate(headers_t1):
    table_t1.cell(0, i).paragraphs[0].text = h

data_t1 = [
    ["乳癌 (BC, group 1)", f"{desc_mw.loc['BC','n']:.0f}", f"{desc_mw.loc['BC','mean']:.2f} ± {desc_mw.loc['BC','std']:.2f}",
     f"{desc_mw.loc['BC','median']:.2f}", f"{desc_mw.loc['BC','iqr']:.2f} [{desc_mw.loc['BC','q25']:.1f}, {desc_mw.loc['BC','q75']:.1f}]",
     f"[{desc_mw.loc['BC','min']:.0f}, {desc_mw.loc['BC','max']:.0f}]", f"W={shapiro_bc.statistic:.3f}, p={shapiro_bc.pvalue:.3f}", f"{mean_rank_bc:.2f}"],
    ["肺腺癌 (LUAD, group 2)", f"{desc_mw.loc['LUAD','n']:.0f}", f"{desc_mw.loc['LUAD','mean']:.2f} ± {desc_mw.loc['LUAD','std']:.2f}",
     f"{desc_mw.loc['LUAD','median']:.2f}", f"{desc_mw.loc['LUAD','iqr']:.2f} [{desc_mw.loc['LUAD','q25']:.1f}, {desc_mw.loc['LUAD','q75']:.1f}]",
     f"[{desc_mw.loc['LUAD','min']:.0f}, {desc_mw.loc['LUAD','max']:.0f}]", f"W={shapiro_luad.statistic:.3f}, p={shapiro_luad.pvalue:.3f}**", f"{mean_rank_luad:.2f}"]
]
for row_idx, r_data in enumerate(data_t1, start=1):
    for col_idx, text in enumerate(r_data):
        table_t1.cell(row_idx, col_idx).paragraphs[0].text = text

style_table(table_t1, [Inches(1.4), Inches(0.7), Inches(1.3), Inches(0.8), Inches(1.2), Inches(1.1), Inches(1.2), Inches(1.0)])

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.25
p.paragraph_format.space_before = Pt(8)
r = p.add_run(
    "1.2 統計檢定方法評估與選擇原因：\n"
    "• 常態性檢定：Shapiro-Wilk 檢定顯示 LUAD 組之 p = 0.0033 < 0.01，顯著違背常態分佈假設；且兩組樣本數均偏小 (n < 30)。\n"
    "• 變異數同質性：Levene 檢定統計量 F = 52.66, p = 2.80×10⁻⁶ < 0.001，顯示兩組變異數嚴重不均齊。\n"
    "• 綜合判定：由於非常態性與變異數異質性，本研究應採用「曼-惠特尼 U 檢定 (Mann-Whitney U Test)」作為最適分析工具。"
)
set_run_font(r, "微軟正黑體", 10.5)

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.25
p.paragraph_format.space_before = Pt(6)
r = p.add_run(
    f"1.3 統計檢定結果與學術結論：\n"
    f"• 秩和與 U 值：BC 組秩和 W_BC = {rank_sum_bc:.1f} (平均秩次 = {mean_rank_bc:.2f})；LUAD 組秩和 W_LUAD = {rank_sum_luad:.1f} (平均秩次 = {mean_rank_luad:.2f})。\n"
    f"• 曼-惠特尼 U 統計量：U_BC = {U_bc:.1f}, U_LUAD = {U_luad:.1f} (SPSS 輸出取較小值 U = {U_min:.1f})，雙尾顯著性 p = {mw_res.pvalue:.4f} < 0.001。\n"
    f"• 效應值 (Effect Size)：Rank-biserial correlation r = {r_biserial:.3f}，標準化常態轉換效應值 r = {effect_r:.3f}，皆屬於極大效應量 (Large effect)。\n"
    f"• 學術結論：乳癌 (BC) 患者腫瘤組織之 HER2 表現量 (中位數 446.5) 極顯著高於肺腺癌 (LUAD) 患者 (中位數 118.0) (U = {U_min:.1f}, p < 0.001)。"
)
set_run_font(r, "微軟正黑體", 10.5)

p_img1 = doc.add_paragraph()
p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img1.paragraph_format.space_before = Pt(8)
p_img1.paragraph_format.space_after = Pt(2)
doc.add_picture(fig1_path, width=Inches(5.0))
p_cap1 = doc.add_paragraph()
p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap1.paragraph_format.space_after = Pt(16)
r_cap1 = p_cap1.add_run("圖 1. 乳癌 (BC) 與肺腺癌 (LUAD) 之 HER2 表現量箱形圖 (Boxplot & Stripplot)")
set_run_font(r_cap1, "微軟正黑體", 9.5, bold=True, color_rgb=(89, 89, 89))

doc.add_page_break()

# 第二部分
h1_3 = doc.add_heading(level=1)
r_h1_3 = h1_3.add_run("第二部分：不同血型 (A, B, O, AB) 族群之身高差異分析")
set_run_font(r_h1_3, "微軟正黑體", 15, bold=True, color_rgb=(31, 78, 121))

p = doc.add_paragraph()
r = p.add_run("2.1 研究背景與資料概況")
set_run_font(r, "微軟正黑體", 12, bold=True, color_rgb=(51, 51, 51))

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.25
r = p.add_run(
    "本資料集 (KW_Blood type.xlsx) 旨在評估個體的 ABO 血型類別是否與身高 (Height, cm) 存在關聯。"
    "樣本共計 N = 44 人，依血型分為 A 型 (n=12)、B 型 (n=10)、O 型 (n=9)、AB 型 (n=13) 四組獨立樣本。"
)
set_run_font(r, "微軟正黑體", 11)

p = doc.add_paragraph()
r = p.add_run("表 2. 各血型組身高之描述性統計與檢定結果")
set_run_font(r, "微軟正黑體", 10.5, bold=True, color_rgb=(31, 78, 121))

table_t2 = doc.add_table(rows=5, cols=8)
headers_t2 = ["血型 (組別)", "樣本數 (n)", "平均值 ± 標準差", "中位數", "四分位距 (IQR)", "極值範圍 [Min, Max]", "常態性檢定 (W, p)", "平均秩次 (Mean Rank)"]
for i, h in enumerate(headers_t2):
    table_t2.cell(0, i).paragraphs[0].text = h

data_t2 = [
    ["A 型 (group 1)", f"{desc_bt.loc['A','n']:.0f}", f"{desc_bt.loc['A','mean']:.2f} ± {desc_bt.loc['A','std']:.2f}",
     f"{desc_bt.loc['A','median']:.2f}", f"{desc_bt.loc['A','iqr']:.2f} [{desc_bt.loc['A','q25']:.1f}, {desc_bt.loc['A','q75']:.1f}]",
     f"[{desc_bt.loc['A','min']:.0f}, {desc_bt.loc['A','max']:.0f}]", f"W={shapiro_bt['A'].statistic:.3f}, p={shapiro_bt['A'].pvalue:.3f}", f"{mean_ranks_bt['A']:.2f}"],
    ["B 型 (group 2)", f"{desc_bt.loc['B','n']:.0f}", f"{desc_bt.loc['B','mean']:.2f} ± {desc_bt.loc['B','std']:.2f}",
     f"{desc_bt.loc['B','median']:.2f}", f"{desc_bt.loc['B','iqr']:.2f} [{desc_bt.loc['B','q25']:.1f}, {desc_bt.loc['B','q75']:.1f}]",
     f"[{desc_bt.loc['B','min']:.0f}, {desc_bt.loc['B','max']:.0f}]", f"W={shapiro_bt['B'].statistic:.3f}, p={shapiro_bt['B'].pvalue:.3f}", f"{mean_ranks_bt['B']:.2f}"],
    ["O 型 (group 3)", f"{desc_bt.loc['O','n']:.0f}", f"{desc_bt.loc['O','mean']:.2f} ± {desc_bt.loc['O','std']:.2f}",
     f"{desc_bt.loc['O','median']:.2f}", f"{desc_bt.loc['O','iqr']:.2f} [{desc_bt.loc['O','q25']:.1f}, {desc_bt.loc['O','q75']:.1f}]",
     f"[{desc_bt.loc['O','min']:.0f}, {desc_bt.loc['O','max']:.0f}]", f"W={shapiro_bt['O'].statistic:.3f}, p={shapiro_bt['O'].pvalue:.3f}", f"{mean_ranks_bt['O']:.2f}"],
    ["AB 型 (group 4)", f"{desc_bt.loc['AB','n']:.0f}", f"{desc_bt.loc['AB','mean']:.2f} ± {desc_bt.loc['AB','std']:.2f}",
     f"{desc_bt.loc['AB','median']:.2f}", f"{desc_bt.loc['AB','iqr']:.2f} [{desc_bt.loc['AB','q25']:.1f}, {desc_bt.loc['AB','q75']:.1f}]",
     f"[{desc_bt.loc['AB','min']:.0f}, {desc_bt.loc['AB','max']:.0f}]", f"W={shapiro_bt['AB'].statistic:.3f}, p={shapiro_bt['AB'].pvalue:.3f}", f"{mean_ranks_bt['AB']:.2f}"]
]
for row_idx, r_data in enumerate(data_t2, start=1):
    for col_idx, text in enumerate(r_data):
        table_t2.cell(row_idx, col_idx).paragraphs[0].text = text

style_table(table_t2, [Inches(1.3), Inches(0.7), Inches(1.3), Inches(0.8), Inches(1.2), Inches(1.1), Inches(1.2), Inches(1.0)])

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.25
p.paragraph_format.space_before = Pt(8)
r = p.add_run(
    f"2.2 統計檢定方法評估與結果：\n"
    f"• 常態性與同質性檢定：四組之 Shapiro-Wilk 常態性檢定 p 值均大於 0.05 (A型 p=0.481, B型 p=0.597, O型 p=0.606, AB型 p=0.970)；"
    f"Levene 變異數同質性檢定 F = {levene_bt.statistic:.3f}, p = {levene_bt.pvalue:.3f} > 0.05，滿足變異數均齊假設。\n"
    f"• Kruskal-Wallis 檢定 (非母數分析)：H 統計量 = {kw_bt.statistic:.3f}, 自由度 df = 3, p = {kw_bt.pvalue:.3f} > 0.05。\n"
    f"• 單因子變異數分析 (One-way ANOVA 母數對照)：F = {anova_bt.statistic:.3f}, df = (3, 40), p = {anova_bt.pvalue:.3f} > 0.05。\n"
    f"• 效應值：Epsilon-squared ε² = {epsilon2_bt:.3f} ≈ 0.000，表示分組對身高無解釋力。\n"
    f"• 學術結論：無論使用非母數 Kruskal-Wallis 檢定或母數 ANOVA 檢定，p 值均遠大於 0.05，無法拒絕虛無假設 H₀。"
    f"研究結果證實：不同血型族群在身高表現上無統計學顯著差異，血型與身高發展無顯著關聯。"
)
set_run_font(r, "微軟正黑體", 10.5)

p_img2 = doc.add_paragraph()
p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img2.paragraph_format.space_before = Pt(8)
p_img2.paragraph_format.space_after = Pt(2)
doc.add_picture(fig2_path, width=Inches(5.0))
p_cap2 = doc.add_paragraph()
p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap2.paragraph_format.space_after = Pt(16)
r_cap2 = p_cap2.add_run("圖 2. 各血型族群之身高分佈箱形圖 (Boxplot & Stripplot)")
set_run_font(r_cap2, "微軟正黑體", 9.5, bold=True, color_rgb=(89, 89, 89))

doc.add_page_break()

# 第三部分
h1_4 = doc.add_heading(level=1)
r_h1_4 = h1_4.add_run("第三部分：四大職業運動聯盟球員薪資與年齡分析")
set_run_font(r_h1_4, "微軟正黑體", 15, bold=True, color_rgb=(31, 78, 121))

p = doc.add_paragraph()
r = p.add_run("3.1 研究背景與資料概況")
set_run_font(r, "微軟正黑體", 12, bold=True, color_rgb=(51, 51, 51))

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.25
r = p.add_run(
    "本資料集 (KW_Pro_salary.xlsx) 納入北美四大職業運動賽事頂尖球員：NBA (男子籃球, n=10)、MLB (男子棒球, n=10)、"
    "NFL (美式足球, n=10) 以及 WNBA (女子籃球, n=10)，總計 N = 40 名運動員。"
    "主要探討不同職業運動聯盟在「年薪 (Annual Salary)」與「年齡 (Age)」上的結構性差異，並兼評男女性別薪資鴻溝。"
)
set_run_font(r, "微軟正黑體", 11)

p = doc.add_paragraph()
r = p.add_run("表 3. 各職業運動聯盟年薪 (美元) 之描述性統計與檢定")
set_run_font(r, "微軟正黑體", 10.5, bold=True, color_rgb=(31, 78, 121))

table_t3 = doc.add_table(rows=5, cols=7)
headers_t3 = ["聯盟 (運動項目)", "樣本數", "平均年薪 (美元)", "中位數年薪 (美元)", "四分位距 (IQR)", "常態檢定 (W, p)", "平均秩次 (Salary Rank)"]
for i, h in enumerate(headers_t3):
    table_t3.cell(0, i).paragraphs[0].text = h

data_t3 = [
    ["NBA (男子籃球)", "10", f"${desc_pro.loc['NBA','mean']:,.0f}", f"${desc_pro.loc['NBA','median']:,.0f}", f"${desc_pro.loc['NBA','iqr']:,.0f}", f"W={shapiro_pro['NBA'].statistic:.3f}, p={shapiro_pro['NBA'].pvalue:.3f}", f"{mean_ranks_pro['NBA']:.2f}"],
    ["MLB (男子棒球)", "10", f"${desc_pro.loc['MLB','mean']:,.0f}", f"${desc_pro.loc['MLB','median']:,.0f}", f"${desc_pro.loc['MLB','iqr']:,.0f}", f"W={shapiro_pro['MLB'].statistic:.3f}, p={shapiro_pro['MLB'].pvalue:.3f}**", f"{mean_ranks_pro['MLB']:.2f}"],
    ["NFL (美式足球)", "10", f"${desc_pro.loc['NFL','mean']:,.0f}", f"${desc_pro.loc['NFL','median']:,.0f}", f"${desc_pro.loc['NFL','iqr']:,.0f}", f"W={shapiro_pro['NFL'].statistic:.3f}, p={shapiro_pro['NFL'].pvalue:.3f}", f"{mean_ranks_pro['NFL']:.2f}"],
    ["WNBA (女子籃球)", "10", f"${desc_pro.loc['WNBA','mean']:,.0f}", f"${desc_pro.loc['WNBA','median']:,.0f}", f"${desc_pro.loc['WNBA','iqr']:,.0f}", f"W={shapiro_pro['WNBA'].statistic:.3f}, p={shapiro_pro['WNBA'].pvalue:.3f}***", f"{mean_ranks_pro['WNBA']:.2f}"]
]
for row_idx, r_data in enumerate(data_t3, start=1):
    for col_idx, text in enumerate(r_data):
        table_t3.cell(row_idx, col_idx).paragraphs[0].text = text

style_table(table_t3, [Inches(1.4), Inches(0.6), Inches(1.4), Inches(1.4), Inches(1.3), Inches(1.2), Inches(1.1)])

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.25
p.paragraph_format.space_before = Pt(8)
r = p.add_run(
    f"3.2 年薪檢定方法與 Kruskal-Wallis 檢定結果：\n"
    f"• 常態性與同質性檢定：MLB (p = 0.0083) 與 WNBA (p = 0.00087) 嚴重違背常態分佈假設；Levene 變異數同質性檢定 F = {levene_pro.statistic:.3f}, p = {levene_pro.pvalue:.3f} < 0.05，變異數顯著不均齊。因此必須採用非母數統計檢定。\n"
    f"• Kruskal-Wallis 檢定結果：H 統計量 = {kw_pro.statistic:.3f}, 自由度 df = 3, p = {kw_pro.pvalue:.2e} < 0.001，達極高度統計顯著差異。\n"
    f"• 效應值：Epsilon-squared ε² = {epsilon2_pro:.3f}，表示高達 68.4% 的薪資秩次變異可由所屬職業運動聯盟解釋 (超大型效應量)。\n"
    f"• 性別差異 Mann-Whitney U 檢定：男運動員 (n=30) vs 女運動員 (n=10) 年薪檢定統計量 U = {mw_gender.statistic:.1f}, p = {mw_gender.pvalue:.2e} < 0.001，兩性薪資存在巨大鴻溝。"
)
set_run_font(r, "微軟正黑體", 10.5)

p = doc.add_paragraph()
r = p.add_run("表 4. 年薪之 Dunn's 事後多重比較 (Dunn's Post-Hoc Test with Bonferroni Correction)")
set_run_font(r, "微軟正黑體", 10.5, bold=True, color_rgb=(31, 78, 121))

table_t4 = doc.add_table(rows=7, cols=6)
headers_t4 = ["兩兩比較分組", "平均秩次差 (Diff)", "標準化 Z 值", "原始 p 值 (Raw p)", "Bonferroni 校正 p 值", "顯著性判定 (α=0.05)"]
for i, h in enumerate(headers_t4):
    table_t4.cell(0, i).paragraphs[0].text = h

for row_idx, r_dict in enumerate(dunn_sal_results, start=1):
    table_t4.cell(row_idx, 0).paragraphs[0].text = r_dict['Comparison']
    table_t4.cell(row_idx, 1).paragraphs[0].text = f"{r_dict['Rank_Diff']:.2f}"
    table_t4.cell(row_idx, 2).paragraphs[0].text = f"{r_dict['Z']:.3f}"
    table_t4.cell(row_idx, 3).paragraphs[0].text = f"{r_dict['p_raw']:.5f}"
    table_t4.cell(row_idx, 4).paragraphs[0].text = f"{r_dict['p_bonf']:.5f}"
    table_t4.cell(row_idx, 5).paragraphs[0].text = "顯著 (***)" if r_dict['p_bonf'] < 0.001 else ("顯著 (**)" if r_dict['p_bonf'] < 0.01 else ("顯著 (*)" if r_dict['p_bonf'] < 0.05 else "無顯著差異 (ns)"))

style_table(table_t4, [Inches(1.5), Inches(1.1), Inches(1.0), Inches(1.1), Inches(1.4), Inches(1.4)])

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.25
p.paragraph_format.space_before = Pt(8)
r = p.add_run(
    "3.3 年薪事後檢定學術解讀：\n"
    "• WNBA 的年薪顯著且大幅度落後於三大男子聯盟：WNBA vs NFL (p_bonf < 0.0001)、WNBA vs NBA (p_bonf = 0.0021)、WNBA vs MLB (p_bonf = 0.0323)。頂尖 WNBA 球員中位數年薪僅約 20 萬美元，而三大男子聯盟頂尖合約均達 3,800 萬至 5,200 萬美元以上。\n"
    "• 三大男子聯盟之間無統計學顯著差異：NBA vs MLB (p_bonf = 1.000)、NBA vs NFL (p_bonf = 0.742)、MLB vs NFL (p_bonf = 0.118)，彼此年薪在頂尖合約層級未達統計顯著差異。"
)
set_run_font(r, "微軟正黑體", 10.5)

p_img3 = doc.add_paragraph()
p_img3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img3.paragraph_format.space_before = Pt(8)
p_img3.paragraph_format.space_after = Pt(2)
doc.add_picture(fig3_path, width=Inches(6.2))
p_cap3 = doc.add_paragraph()
p_cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap3.paragraph_format.space_after = Pt(16)
r_cap3 = p_cap3.add_run("圖 3. 四大職業聯盟頂尖球員年薪比較 (左：線性百萬美元尺度；右：對數尺度與事後檢定標記)")
set_run_font(r_cap3, "微軟正黑體", 9.5, bold=True, color_rgb=(89, 89, 89))

p = doc.add_paragraph()
r = p.add_run("3.4 球員年齡 (Age) 之 Kruskal-Wallis 檢定與事後分析")
set_run_font(r, "微軟正黑體", 12, bold=True, color_rgb=(51, 51, 51))

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.25
r = p.add_run(
    f"• 描述統計：MLB 平均年齡最高 (33.6 ± 4.58 歲, 中位數 33.5 歲)，次為 NBA (30.2 ± 5.55 歲, 中位數 30.0 歲)，"
    f"WNBA (29.8 ± 4.98 歲, 中位數 29.0 歲)，NFL 最年輕 (27.6 ± 1.96 歲, 中位數 27.5 歲)。\n"
    f"• 統計檢定：Kruskal-Wallis H 統計量 = {kw_age.statistic:.3f}, df = 3, p = {kw_age.pvalue:.4f} < 0.05 (* 達顯著差異)。\n"
    f"• Dunn 事後檢定：僅「MLB vs NFL」比較達統計顯著差異 (Z = 2.917, p_raw = 0.0035, p_bonf = 0.0212 < 0.05)，棒球巨星年齡顯著高於美式足球四分衛；其餘各組兩兩比較校正後均無顯著差異 (p_bonf > 0.05)。"
)
set_run_font(r, "微軟正黑體", 10.5)

p_img4 = doc.add_paragraph()
p_img4.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img4.paragraph_format.space_before = Pt(8)
p_img4.paragraph_format.space_after = Pt(2)
doc.add_picture(fig4_path, width=Inches(5.0))
p_cap4 = doc.add_paragraph()
p_cap4.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap4.paragraph_format.space_after = Pt(16)
r_cap4 = p_cap4.add_run("圖 4. 四大職業聯盟頂尖球員年齡分佈箱形圖 (Boxplot & Stripplot)")
set_run_font(r_cap4, "微軟正黑體", 9.5, bold=True, color_rgb=(89, 89, 89))

doc.add_page_break()

# 第四部分
h1_5 = doc.add_heading(level=1)
r_h1_5 = h1_5.add_run("第四部分：統計方法學總結與 SPSS 實務對照手冊")
set_run_font(r_h1_5, "微軟正黑體", 15, bold=True, color_rgb=(31, 78, 121))

p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.3
r = p.add_run(
    "4.1 何時採用非母數統計 (Non-parametric Statistics)？\n"
    "在生物統計與醫學研究中，母數檢定 (如獨立雙樣本 t 檢定、One-way ANOVA) 依賴三大嚴格假設：(1) 常態分佈 (Normality)、(2) 變異數同質性 (Homoscedasticity)、(3) 樣本獨立性。當遇到以下情境時，非母數檢定為必然選擇：\n"
    "1. 樣本數過小 (n < 30) 且常態性檢定 (Shapiro-Wilk) 達顯著違背常態 (p < 0.05)，如本報告之 LUAD 組 HER2 表現量、WNBA/MLB 年薪。\n"
    "2. 變異數嚴重不均齊 (Levene's test p < 0.05)，極端值對平均值與變異數造成過度扭曲。\n"
    "3. 資料為等級資料 (Ordinal) 或具有明顯偏態分佈。\n\n"
    "4.2 曼-惠特尼 U 檢定 (Mann-Whitney U Test) 原理：\n"
    "• 適用於兩組獨立樣本。將所有樣本合併由小到大排序並賦予秩次 (Rank, 遇同分取平均秩次)，再分別加總兩組秩次得到 R₁ 與 R₂。\n"
    "• 計算公式：U₁ = R₁ - n₁(n₁+1)/2，U₂ = R₂ - n₂(n₂+1)/2。檢定統計量 U 滿足 U₁ + U₂ = n₁n₂。\n"
    "• SPSS 報告規範：SPSS 輸出之 Mann-Whitney U 通常報告 Min(U₁, U₂)，Wilcoxon W 報告較小樣本組之秩和。研究者可報告 U、Z 值與對應 p 值。\n\n"
    "4.3 克魯斯卡-沃利斯 H 檢定 (Kruskal-Wallis H Test) 原理：\n"
    "• 適用於三組或以上獨立樣本 (單因子變異數分析之非母數替代法)。將全部 N 個觀察值混合排序編秩，計算各組平均秩次。\n"
    "• 統計量 H 近似於自由度為 k-1 之卡方 (Chi-square) 分佈。若 H 檢定達顯著水準 (p < 0.05)，僅代表「至少有兩組之間存在差異」，必須進一步執行 Dunn's 事後檢定並搭配 Bonferroni 或 Holm 多重比較校正，以避免第一型錯誤 (Type I Error) 膨脹。\n\n"
    "4.4 SPSS 實務操作步驟指引：\n"
    "• Mann-Whitney U 檢定操作：\n"
    "  分析 (Analyze) -> 無母數檢定 (Nonparametric Tests) -> 舊式對話框 (Legacy Dialogs) -> 2 個獨立樣本 (2 Independent Samples) -> 檢定變數放入應變數，分組變數放入分組代碼 (定義群組 1, 2) -> 勾選 Mann-Whitney U。\n"
    "• Kruskal-Wallis 檢定操作：\n"
    "  分析 (Analyze) -> 無母數檢定 (Nonparametric Tests) -> 舊式對話框 (Legacy Dialogs) -> K 個獨立樣本 (K Independent Samples) -> 檢定變數放入應變數，分組變數放入分組欄位 (定義範圍 1 到 k) -> 勾選 Kruskal-Wallis H。"
)
set_run_font(r, "微軟正黑體", 10.5)

word_path_chi = os.path.join(base_dir, "生統實習_三份資料統計分析報告.docx")
word_path_eng = os.path.join(base_dir, "Biostatistics_Statistical_Analysis_Report.docx")
doc.save(word_path_chi)
doc.save(word_path_eng)
print(f"已儲存中文報告: {word_path_chi}")
print(f"已儲存英文報告: {word_path_eng}")
print("全部運算與報表生成完成！")
