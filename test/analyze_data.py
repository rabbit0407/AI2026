import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
from scipy import stats

# Set Korean font for matplotlib (Windows default)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# Paths
data_dir = r"c:\Users\user\Desktop\test\data"
file1 = os.path.join(data_dir, "삶의_만족도_시도__20260606195059.xlsx")
file2 = os.path.join(data_dir, "인구십만명당_자살률_시도_시_군_구__20260606194913.xlsx")

artifact_dir = r"C:\Users\user\.gemini\antigravity\brain\927c726d-a694-4c34-98ef-c0d863dd586a"
scratch_dir = os.path.join(artifact_dir, "scratch")
os.makedirs(scratch_dir, exist_ok=True)

# 1. Load Data
df1_raw = pd.read_excel(file1, sheet_name="데이터")
df2_raw = pd.read_excel(file2, sheet_name="데이터")

# 2. Process File 1 (Life Satisfaction)
df1 = df1_raw.copy()
df1['행정구역별(1)'] = df1['행정구역별(1)'].ffill()
df1['특성별(1)'] = df1['특성별(1)'].ffill()

# Drop rows that are headers
df1 = df1[df1['행정구역별(1)'] != '행정구역별(1)']

# Helper to normalize region names
def normalize_region(name):
    if not isinstance(name, str):
        return name
    name = name.strip()
    if name in ['전북특별자치도', '전라북도']:
        return '전북특별자치도'
    if name in ['제주특별자치도', '제주도']:
        return '제주특별자치도'
    return name

df1['행정구역별(1)'] = df1['행정구역별(1)'].apply(normalize_region)

# Map Gender
# 전체/계 -> Total, 성별/남자 -> Male, 성별/여자 -> Female
def get_gender_label(row):
    c1 = row['특성별(1)'].strip()
    c2 = row['특성별(2)'].strip()
    if c1 == '전체' and c2 == '계':
        return 'Total'
    elif c1 == '성별' and c2 == '남자':
        return 'Male'
    elif c1 == '성별' and c2 == '여자':
        return 'Female'
    return None

df1['Gender'] = df1.apply(get_gender_label, axis=1)
df1 = df1.dropna(subset=['Gender'])

# Melt and reconstruct File 1 data year by year
years = ['2020', '2021', '2022', '2023', '2024']
life_sat_list = []

for idx, row in df1.iterrows():
    region = row['행정구역별(1)']
    gender = row['Gender']
    
    for year in years:
        # Columns in Excel:
        # year: 계 (100)
        # year.1: 매우 만족
        # year.2: 약간 만족
        # year.3: 보통
        # year.4: 약간 불만족
        # year.5: 매우 불만족
        try:
            col_very_sat = f"{year}.1"
            col_somewh_sat = f"{year}.2"
            col_neutral = f"{year}.3"
            col_somewh_dissat = f"{year}.4"
            col_very_dissat = f"{year}.5"
            
            # Convert values to float
            val_very_sat = pd.to_numeric(row[col_very_sat], errors='coerce')
            val_somewh_sat = pd.to_numeric(row[col_somewh_sat], errors='coerce')
            val_neutral = pd.to_numeric(row[col_neutral], errors='coerce')
            val_somewh_dissat = pd.to_numeric(row[col_somewh_dissat], errors='coerce')
            val_very_dissat = pd.to_numeric(row[col_very_dissat], errors='coerce')
            
            # Calculate metrics
            satisfied_pct = val_very_sat + val_somewh_sat
            dissatisfied_pct = val_somewh_dissat + val_very_dissat
            
            # Average score (weighted: very_sat=5, somewh_sat=4, neutral=3, somewh_dissat=2, very_dissat=1)
            # sum of percentages might be around 100.
            total_valid_pct = val_very_sat + val_somewh_sat + val_neutral + val_somewh_dissat + val_very_dissat
            if total_valid_pct > 0:
                avg_score = (val_very_sat * 5 + val_somewh_sat * 4 + val_neutral * 3 + val_somewh_dissat * 2 + val_very_dissat * 1) / total_valid_pct
            else:
                avg_score = np.nan
                
            life_sat_list.append({
                'Region': region,
                'Gender': gender,
                'Year': int(year),
                'Very_Satisfied_Pct': val_very_sat,
                'Somewhat_Satisfied_Pct': val_somewh_sat,
                'Neutral_Pct': val_neutral,
                'Somewhat_Dissatisfied_Pct': val_somewh_dissat,
                'Very_Dissatisfied_Pct': val_very_dissat,
                'Satisfied_Pct': satisfied_pct,
                'Dissatisfied_Pct': dissatisfied_pct,
                'Satisfaction_Score': avg_score
            })
        except Exception as e:
            # Column might not exist or some other issue
            print(f"Skipping year {year} for {region} / {gender}: {e}")

df_life_sat = pd.DataFrame(life_sat_list)

# 3. Process File 2 (Suicide Rate)
df2 = df2_raw.copy()
df2 = df2[df2['행정구역별(1)'] != '행정구역별(1)']
df2['행정구역별(1)'] = df2['행정구역별(1)'].apply(normalize_region)

suicide_list = []
for idx, row in df2.iterrows():
    region = row['행정구역별(1)']
    for year in years:
        try:
            # Columns in Excel:
            # year: 계
            # year.1: 남자
            # year.2: 여자
            col_total = f"{year}"
            col_male = f"{year}.1"
            col_female = f"{year}.2"
            
            val_total = pd.to_numeric(row[col_total], errors='coerce')
            val_male = pd.to_numeric(row[col_male], errors='coerce')
            val_female = pd.to_numeric(row[col_female], errors='coerce')
            
            suicide_list.append({'Region': region, 'Year': int(year), 'Gender': 'Total', 'Suicide_Rate': val_total})
            suicide_list.append({'Region': region, 'Year': int(year), 'Gender': 'Male', 'Suicide_Rate': val_male})
            suicide_list.append({'Region': region, 'Year': int(year), 'Gender': 'Female', 'Suicide_Rate': val_female})
        except Exception as e:
            print(f"Skipping year {year} for suicide rate in {region}: {e}")

df_suicide = pd.DataFrame(suicide_list)

# 4. Merge Datasets
df_merged = pd.merge(df_life_sat, df_suicide, on=['Region', 'Year', 'Gender'], how='inner')

# Drop national total row from region list for region-specific correlation analysis, but keep it in a separate df if needed
df_regions_only = df_merged[df_merged['Region'] != '전국'].copy()
df_national_only = df_merged[df_merged['Region'] == '전국'].copy()

# Save merged dataset
df_merged.to_csv(os.path.join(scratch_dir, "merged_data.csv"), index=False, encoding='utf-8-sig')

# 5. Correlation Analysis
results_txt = []
results_txt.append("==================================================")
results_txt.append("      LIFE SATISFACTION VS SUICIDE RATE CORRELATION")
results_txt.append("==================================================")

# Drop missing values for correlation
df_regions_clean = df_regions_only.dropna(subset=['Satisfied_Pct', 'Dissatisfied_Pct', 'Satisfaction_Score', 'Suicide_Rate'])

results_txt.append(f"Total data points (excl. National Total): {len(df_regions_clean)}")

# Helper to format correlation output
def get_corr_summary(df, x_col, y_col):
    if len(df) < 3:
        return "Insufficient data"
    pearson_r, pearson_p = stats.pearsonr(df[x_col], df[y_col])
    spearman_r, spearman_p = stats.spearmanr(df[x_col], df[y_col])
    return (f"Pearson r = {pearson_r:.4f} (p-value = {pearson_p:.4e})\n"
            f"Spearman r = {spearman_r:.4f} (p-value = {spearman_p:.4e})")

# Overall Correlation (Pooled over all years and regions)
results_txt.append("\n--- POOLED CORRELATION (ALL YEARS & ALL REGIONS) ---")
for gender in ['Total', 'Male', 'Female']:
    sub = df_regions_clean[df_regions_clean['Gender'] == gender]
    results_txt.append(f"\n[Gender: {gender}] (N={len(sub)})")
    results_txt.append(f"1. Satisfied % vs Suicide Rate:")
    results_txt.append(get_corr_summary(sub, 'Satisfied_Pct', 'Suicide_Rate'))
    results_txt.append(f"2. Dissatisfied % vs Suicide Rate:")
    results_txt.append(get_corr_summary(sub, 'Dissatisfied_Pct', 'Suicide_Rate'))
    results_txt.append(f"3. Avg Satisfaction Score vs Suicide Rate:")
    results_txt.append(get_corr_summary(sub, 'Satisfaction_Score', 'Suicide_Rate'))

# Year-by-Year Correlation (Total Gender)
results_txt.append("\n--- YEAR-BY-YEAR CORRELATION (GENDER: TOTAL) ---")
for yr in sorted(df_regions_clean['Year'].unique()):
    sub = df_regions_clean[(df_regions_clean['Gender'] == 'Total') & (df_regions_clean['Year'] == yr)]
    results_txt.append(f"\n[Year: {yr}] (N={len(sub)})")
    results_txt.append(f"Avg Satisfaction Score vs Suicide Rate:")
    results_txt.append(get_corr_summary(sub, 'Satisfaction_Score', 'Suicide_Rate'))
    results_txt.append(f"Satisfied % vs Suicide Rate:")
    results_txt.append(get_corr_summary(sub, 'Satisfied_Pct', 'Suicide_Rate'))

# Save results text
with open(os.path.join(scratch_dir, "correlation_results.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(results_txt))

# 6. Generate Visualizations
# Plot 1: Scatter plot of Avg Satisfaction Score vs Suicide Rate (Total, Male, Female)
plt.figure(figsize=(10, 6), dpi=150)
colors = {'Total': '#3182bd', 'Male': '#e6550d', 'Female': '#31a354'}
markers = {'Total': 'o', 'Male': 's', 'Female': '^'}

for gender in ['Total', 'Male', 'Female']:
    sub = df_regions_clean[df_regions_clean['Gender'] == gender]
    # Scatter plot
    plt.scatter(sub['Satisfaction_Score'], sub['Suicide_Rate'], 
                color=colors[gender], marker=markers[gender], alpha=0.7, label=f'{gender}', s=50)
    
    # Regression line
    if len(sub) > 1:
        slope, intercept, r_value, p_value, std_err = stats.linregress(sub['Satisfaction_Score'], sub['Suicide_Rate'])
        x_vals = np.linspace(sub['Satisfaction_Score'].min(), sub['Satisfaction_Score'].max(), 100)
        y_vals = slope * x_vals + intercept
        plt.plot(x_vals, y_vals, color=colors[gender], linestyle='--', alpha=0.6,
                 label=f'{gender} Trend (r={r_value:.2f})')

plt.title('삶의 만족도 점수 vs 10만 명당 자살률 상관관계 (2020-2024)', fontsize=14, fontweight='bold', pad=15)
plt.xlabel('삶의 만족도 점수 (5점 만점 환산)', fontsize=12)
plt.ylabel('인구 10만 명당 자살률 (명)', fontsize=12)
plt.grid(True, linestyle=':', alpha=0.6)
plt.legend(frameon=True, facecolor='white', edgecolor='none', shadow=True)
plt.tight_layout()
plt.savefig(os.path.join(artifact_dir, "satisfaction_vs_suicide_scatter.png"), dpi=300)
plt.close()

# Plot 2: Scatter plot specifically for Total, showing region names for the most recent year (2024)
plt.figure(figsize=(10, 7), dpi=150)
sub_2024 = df_regions_clean[(df_regions_clean['Gender'] == 'Total') & (df_regions_clean['Year'] == 2024)]

plt.scatter(sub_2024['Satisfaction_Score'], sub_2024['Suicide_Rate'], color='#1f77b4', s=100, alpha=0.8, edgecolors='black', label='2024년 시도별 데이터')

# Fit line
slope, intercept, r_value, p_value, std_err = stats.linregress(sub_2024['Satisfaction_Score'], sub_2024['Suicide_Rate'])
x_vals = np.linspace(sub_2024['Satisfaction_Score'].min(), sub_2024['Satisfaction_Score'].max(), 100)
y_vals = slope * x_vals + intercept
plt.plot(x_vals, y_vals, color='#ff7f0e', linestyle='-', linewidth=2, label=f'선형 추세선 (r={r_value:.2f}, p={p_value:.3f})')

# Annotate points
for idx, row in sub_2024.iterrows():
    plt.annotate(row['Region'], (row['Satisfaction_Score'], row['Suicide_Rate']),
                 textcoords="offset points", xytext=(0,10), ha='center', fontsize=9, fontweight='bold')

plt.title('2024년 시도별 삶의 만족도 점수 vs 자살률 분포', fontsize=14, fontweight='bold', pad=15)
plt.xlabel('삶의 만족도 점수 (5점 만점)', fontsize=12)
plt.ylabel('인구 10만 명당 자살률 (명)', fontsize=12)
plt.grid(True, linestyle=':', alpha=0.6)
plt.legend(frameon=True, loc='upper right')
plt.tight_layout()
plt.savefig(os.path.join(artifact_dir, "satisfaction_vs_suicide_2024_annotated.png"), dpi=300)
plt.close()

# Plot 3: Trend lines over time for National Average
national_total = df_national_only[df_national_only['Gender'] == 'Total'].sort_values('Year')

fig, ax1 = plt.subplots(figsize=(10, 5), dpi=150)

color = '#1f77b4'
ax1.set_xlabel('연도 (Year)', fontsize=12)
ax1.set_ylabel('전국 평균 삶의 만족도 점수 (5점 만점)', color=color, fontsize=12)
line1 = ax1.plot(national_total['Year'], national_total['Satisfaction_Score'], color=color, marker='o', linewidth=2, label='삶의 만족도 점수')
ax1.tick_params(axis='y', labelcolor=color)
ax1.set_xticks(national_total['Year'])

ax2 = ax1.twinx()  
color = '#d62728'
ax2.set_ylabel('전국 평균 인구 10만 명당 자살률 (명)', color=color, fontsize=12)
line2 = ax2.plot(national_total['Year'], national_total['Suicide_Rate'], color=color, marker='s', linewidth=2, label='자살률')
ax2.tick_params(axis='y', labelcolor=color)

# added these lines
lines = line1 + line2
labels = [l.get_label() for l in lines]
ax1.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=2)

plt.title('연도별 전국 평균 삶의 만족도 및 자살률 추이 (2020-2024)', fontsize=14, fontweight='bold', pad=15)
plt.tight_layout()
plt.savefig(os.path.join(artifact_dir, "national_trends.png"), dpi=300)
plt.close()

# Plot 4: Correlation Matrix Heatmap (using matplotlib imshow)
corr_vars = ['Very_Satisfied_Pct', 'Somewhat_Satisfied_Pct', 'Neutral_Pct', 'Somewhat_Dissatisfied_Pct', 'Very_Dissatisfied_Pct', 'Satisfaction_Score', 'Suicide_Rate']
corr_labels = ['매우 만족 (%)', '약간 만족 (%)', '보통 (%)', '약간 불만족 (%)', '매우 불만족 (%)', '만족도 점수', '자살률']
df_corr_sub = df_regions_clean[df_regions_clean['Gender'] == 'Total'][corr_vars].corr()

plt.figure(figsize=(9, 7), dpi=150)
im = plt.imshow(df_corr_sub.values, cmap='RdYlBu_r', vmin=-1, vmax=1)
plt.colorbar(im)
plt.xticks(range(len(corr_vars)), corr_labels, rotation=45, ha='right')
plt.yticks(range(len(corr_vars)), corr_labels)

# Add numeric annotations
for i in range(len(corr_vars)):
    for j in range(len(corr_vars)):
        val = df_corr_sub.values[i, j]
        plt.text(j, i, f"{val:.2f}", ha='center', va='center', 
                 color='black' if abs(val) < 0.6 else 'white', fontweight='bold')
        
plt.title('삶의 세부 만족도 항목 및 자살률 간의 상관관계 열지도(Heatmap)', fontsize=14, fontweight='bold', pad=15)
plt.tight_layout()
plt.savefig(os.path.join(artifact_dir, "correlation_heatmap.png"), dpi=300)
plt.close()

# Plot 5: 2024 Regional Suicide Rate vs. Satisfaction Comparison (sorted by Suicide Rate)
sub_2024_sorted = sub_2024.sort_values('Suicide_Rate', ascending=False)

fig, ax1 = plt.subplots(figsize=(12, 6), dpi=150)

color_bar = '#d95f02'
ax1.set_ylabel('인구 10만 명당 자살률 (명)', color=color_bar, fontsize=12, fontweight='bold')
bars = ax1.bar(sub_2024_sorted['Region'], sub_2024_sorted['Suicide_Rate'], color=color_bar, alpha=0.7, label='자살률')
ax1.tick_params(axis='y', labelcolor=color_bar)
# Rotate x-axis labels
ax1.set_xticks(range(len(sub_2024_sorted['Region'])))
ax1.set_xticklabels(sub_2024_sorted['Region'], rotation=45, ha='right', fontsize=10, fontweight='bold')

# Add values on top of bars
for bar in bars:
    height = bar.get_height()
    ax1.annotate(f'{height:.1f}',
                 xy=(bar.get_x() + bar.get_width() / 2, height),
                 xytext=(0, 3),  # 3 points vertical offset
                 textcoords="offset points",
                 ha='center', va='bottom', fontsize=8, fontweight='bold')

ax2 = ax1.twinx()
color_line = '#7570b3'
ax2.set_ylabel('삶의 만족도 점수 (5점 만점)', color=color_line, fontsize=12, fontweight='bold')
line = ax2.plot(sub_2024_sorted['Region'], sub_2024_sorted['Satisfaction_Score'], color=color_line, marker='o', linewidth=2, label='삶의 만족도 점수')
ax2.tick_params(axis='y', labelcolor=color_line)

# Combined legend
lines, labels = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines + lines2, labels + labels2, loc='upper right')

plt.title('2024년 시도별 자살률 및 삶의 만족도 비교 (자살률 내림차순)', fontsize=14, fontweight='bold', pad=15)
fig.tight_layout()
plt.savefig(os.path.join(artifact_dir, "regional_comparison_bar.png"), dpi=300)
plt.close()

# Plot 6: Gender Disparity in Suicide Rate and Satisfaction (2020-2024)
national_male = df_national_only[df_national_only['Gender'] == 'Male'].sort_values('Year')
national_female = df_national_only[df_national_only['Gender'] == 'Female'].sort_values('Year')

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6), dpi=150)

# Left subplot: Suicide Rate
ax1.plot(national_male['Year'], national_male['Suicide_Rate'], color='#e6550d', marker='s', linewidth=2.5, label='남성')
ax1.plot(national_female['Year'], national_female['Suicide_Rate'], color='#31a354', marker='^', linewidth=2.5, label='여성')
ax1.set_title('성별 자살률 추이 비교', fontsize=12, fontweight='bold')
ax1.set_xlabel('연도')
ax1.set_ylabel('인구 10만 명당 자살률 (명)')
ax1.set_xticks(national_male['Year'])
ax1.grid(True, linestyle=':', alpha=0.6)
ax1.legend()

# Right subplot: Satisfaction Score
ax2.plot(national_male['Year'], national_male['Satisfaction_Score'], color='#e6550d', linestyle='--', marker='s', linewidth=2, label='남성')
ax2.plot(national_female['Year'], national_female['Satisfaction_Score'], color='#31a354', linestyle='--', marker='^', linewidth=2, label='여성')
ax2.set_title('성별 삶의 만족도 점수 추이 비교', fontsize=12, fontweight='bold')
ax2.set_xlabel('연도')
ax2.set_ylabel('삶의 만족도 점수 (5점 만점)')
ax2.set_xticks(national_male['Year'])
ax2.grid(True, linestyle=':', alpha=0.6)
ax2.legend()

plt.suptitle('전국 성별 자살률 및 삶의 만족도 격차 분석 (2020-2024)', fontsize=14, fontweight='bold', y=0.98)
plt.tight_layout()
plt.savefig(os.path.join(artifact_dir, "gender_disparity_trends.png"), dpi=300)
plt.close()

print("Analysis and visualizations completed successfully!")
