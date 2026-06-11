import os
from fpdf import FPDF
from fpdf.enums import XPos, YPos

# Paths
project_dir = r"c:\Users\user\Desktop\test"
pdf_path = os.path.join(project_dir, "삶의_만족도와_자살률_상관관계_분석_보고서.pdf")
images_dir = os.path.join(project_dir, "images")

class PDFReport(FPDF):
    def header(self):
        pass

    def footer(self):
        self.set_y(-15)
        self.set_font("Malgun", "", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"페이지 {self.page_no()}/{{nb}}", align="C")

# Initialize PDF
pdf = PDFReport(orientation="P", unit="mm", format="A4")
pdf.alias_nb_pages()

# Load Korean Fonts from Windows Fonts folder
font_regular = r"C:\Windows\Fonts\malgun.ttf"
font_bold = r"C:\Windows\Fonts\malgunbd.ttf"

if not os.path.exists(font_regular):
    font_regular = "Helvetica"
    font_bold = "Helvetica-Bold"
    print("Warning: Malgun Gothic font not found. Falling back to Helvetica.")
else:
    pdf.add_font("Malgun", "", font_regular)
    pdf.add_font("Malgun", "B", font_bold)

pdf.set_margins(20, 20, 20)
pdf.add_page()

# Helper function to print title
def print_title():
    pdf.set_font("Malgun", "B", 18)
    pdf.set_text_color(31, 78, 120)  # Primary Deep Blue
    pdf.cell(0, 10, "삶의 만족도와 자살률의 상관관계 분석 보고서", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 10, "(2020 ~ 2024)", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_font("Malgun", "", 11)
    pdf.set_text_color(89, 89, 89)   # Gray
    pdf.cell(0, 8, "KOSIS 사회조사 및 사망원인통계 데이터 기반 분석", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(10)

# Helper function for headings
def print_heading(text):
    pdf.ln(5)
    pdf.set_font("Malgun", "B", 13)
    pdf.set_text_color(31, 78, 120)
    pdf.cell(0, 8, text, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_draw_color(31, 78, 120)
    pdf.set_line_width(0.5)
    pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 170, pdf.get_y())
    pdf.ln(4)

# Helper for paragraphs
def print_paragraph(text, bold_prefix="", indent=False):
    pdf.set_font("Malgun", "", 9.5)
    pdf.set_text_color(0, 0, 0)
    
    if bold_prefix:
        pdf.set_font("Malgun", "B", 9.5)
        pdf.write(5, bold_prefix)
        pdf.set_font("Malgun", "", 9.5)
        pdf.write(5, text + "\n")
    else:
        if indent:
            pdf.write(5, "    " + text + "\n")
        else:
            pdf.write(5, text + "\n")

# Start building document
print_title()

# 1. 요약 및 핵심 통찰
print_heading("1. 요약 및 핵심 통찰 (Executive Summary)")
bullet_points = [
    ("• 상관관계의 비유의성 (전체 통합): ", "2020년부터 2024년까지의 5개년 전체 시도별 데이터를 통합하여 분석한 결과, 삶의 만족도 통합 점수와 자살률 간의 단순 선형 상관관계는 통계적으로 유의하지 않은 수준으로 나타났습니다."),
    ("• 코로나19 시기(2020년)의 음의 상관관계: ", "팬데믹 초기인 2020년에는 삶의 만족도가 높은 지역일수록 자살률이 낮아지는 유의미한 음(-)의 상관관계(Pearson r = -0.50, p < 0.05)가 관찰되었으나, 이후 연도에서는 이러한 관계가 해체되었습니다."),
    ("• 세부 설문 항목 분석 (보통 응답의 역설): ", "만족 여부를 묻는 세부 비율 분석 결과, 매우 만족하거나 매우 불만족하는 극단적 응답이 높은 지역보다 \"보통(Neutral)\"이라고 답한 비율이 높은 지역일수록 자살률이 증가하는 뚜렷한 양(+)의 상관관계(Pearson r = 0.41)가 관찰되었습니다. 이는 정서적 무관심이나 미온적 상태가 지역 자살 위험과 깊게 연관되어 있을 수 있음을 시사합니다."),
    ("• 자살률-만족도의 역설 (Daly's Paradox): ", "2024년 시도별 분석 결과, 삶의 만족도가 전국 최상위권인 충청남도와 제주특별자치도에서 오히려 자살률이 각각 2위(34.8명)와 1위(36.3명)를 기록하여 '행복한 지역의 높은 자살률 역설'이 강하게 확인되었습니다."),
    ("• 성별 불일치 (Gender Disparity): ", "남성과 여성의 삶의 만족도는 거의 동일하나, 자살률은 남성이 여성보다 약 2.5배 높았으며 최근 남성 자살률의 증가세가 극도로 가파릅니다.")
]
for prefix, body in bullet_points:
    print_paragraph(body, bold_prefix=prefix)
pdf.ln(5)

# 2. 데이터 개요 및 지표 산출
print_heading("2. 데이터 개요 및 지표 산출")
print_paragraph("본 분석에서는 통계청 국가통계포털(KOSIS)의 두 가지 원천 통계 데이터를 결합하였습니다.")
print_paragraph("1. 삶의 만족도 데이터: 사회조사 기반의 주관적 정서 지표 (매우 만족, 약간 만족, 보통, 약간 불만족, 매우 불만족 비율)")
print_paragraph("2. 자살률 데이터: 사망원인통계 기반의 인구 10만 명당 자살률")
pdf.ln(2)
print_paragraph("정량적인 비교 분석을 위해, 사회조사의 5가지 만족도 설문 답변 비율을 반영한 5점 만점의 '삶의 만족도 점수(Satisfaction Score)'를 아래 수식으로 산출하여 활용했습니다.")
pdf.ln(3)

# Formula box
pdf.set_fill_color(242, 242, 242)
pdf.set_draw_color(200, 200, 200)
pdf.set_font("Malgun", "B", 9)
formula_text = "만족도 점수 = [ (매우 만족 * 5) + (약간 만족 * 4) + (보통 * 3) + (약간 불만족 * 2) + (매우 불만족 * 1) ] / 100"
pdf.cell(0, 10, formula_text, border=1, align="C", fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.ln(5)

# 3. 전국 평균 추이 및 성별 격차
print_heading("3. 전국 평균 추이 및 성별 격차 (National & Gender Trends)")
print_paragraph("지난 5년간 전국 평균 만족도 지표와 자살률의 시계열 추이를 성별로 세분화하여 비교했습니다.")
pdf.ln(2)

# Table 1: National Trends
national_headers = ["성별", "연도", "만족 비율 (%)", "불만족 비율 (%)", "만족도 점수", "자살률 (명)"]
national_rows = [
    ["전체", "2020", "42.7", "12.5", "3.42", "25.7"],
    ["", "2021", "34.0", "22.9", "3.16", "26.0"],
    ["", "2022", "43.3", "14.1", "3.39", "25.2"],
    ["", "2023", "42.2", "14.5", "3.36", "27.3"],
    ["", "2024", "40.1", "12.7", "3.37", "29.1"],
    ["남성", "2020", "43.5", "11.8", "3.45", "35.5"],
    ["", "2021", "33.8", "23.5", "3.15", "35.9"],
    ["", "2022", "43.4", "14.1", "3.39", "35.3"],
    ["", "2023", "42.5", "14.6", "3.37", "38.3"],
    ["", "2024", "40.5", "12.4", "3.38", "41.8"],
    ["여성", "2020", "41.9", "13.0", "3.40", "15.9"],
    ["", "2021", "34.3", "22.4", "3.17", "16.2"],
    ["", "2022", "43.2", "14.0", "3.39", "15.1"],
    ["", "2023", "42.0", "14.5", "3.36", "16.5"],
    ["", "2024", "39.7", "13.1", "3.35", "16.6"]
]

pdf.set_font("Malgun", "", 8.5)
with pdf.table(col_widths=(20, 20, 35, 35, 30, 30), text_align="CENTER") as table:
    row = table.row()
    pdf.set_font("Malgun", "B", 8.5)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(31, 78, 120)
    for h in national_headers:
        row.cell(h)
    
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Malgun", "", 8)
    for r_idx, r_data in enumerate(national_rows):
        row = table.row()
        if r_idx % 2 == 1:
            pdf.set_fill_color(242, 242, 242)
        else:
            pdf.set_fill_color(255, 255, 255)
        for cell_data in r_data:
            row.cell(cell_data)

pdf.ln(5)

# Page Break for images
pdf.add_page()

# Image 1
img_path1 = os.path.join(images_dir, "national_trends.png")
if os.path.exists(img_path1):
    pdf.image(img_path1, x=20, y=20, w=170)
    pdf.set_y(105)
    pdf.set_font("Malgun", "", 8.5)
    pdf.set_text_color(89, 89, 89)
    pdf.cell(0, 5, "<그림 1> 연도별 전국 평균 삶의 만족도 및 자살률 추이", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(10)

# Image 2
img_path2 = os.path.join(images_dir, "gender_disparity_trends.png")
if os.path.exists(img_path2):
    pdf.image(img_path2, x=20, y=120, w=170)
    pdf.set_y(205)
    pdf.set_font("Malgun", "", 8.5)
    pdf.set_text_color(89, 89, 89)
    pdf.cell(0, 5, "<그림 2> 전국 성별 자살률 및 삶의 만족도 격차 비교", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# New Page for Section 4
pdf.add_page()

# 4. 상관관계 분석 결과
print_heading("4. 상관관계 분석 결과 (Correlation Analysis)")
print_paragraph("전국 통계를 제외하고 17개 시도별로 데이터를 매칭하여 상관관계 통계 분석을 수행했습니다.")
print_paragraph("아래 표 2는 5개년 전체 기간과 전 지역 데이터를 병합(N=85)하여 계산한 전체 통합 상관관계 결과입니다.")
pdf.ln(2)

# Table 2: Pooled Correlation
table2_headers = ["성별", "독립 변수 (X)", "종속 변수 (Y)", "피어슨 r", "p-value", "스피어먼 ρ", "p-value"]
table2_rows = [
    ["전체", "만족 비율 (%)", "자살률", "-0.1704", "1.19e-01", "-0.1359", "2.15e-01"],
    ["", "불만족 비율 (%)", "자살률", "-0.1767", "1.06e-01", "-0.1685", "1.23e-01"],
    ["", "만족도 점수", "자살률", "-0.0452", "6.81e-01", "-0.0192", "8.61e-01"],
    ["남성", "만족 비율 (%)", "자살률", "-0.1208", "2.71e-01", "-0.0158", "8.86e-01"],
    ["", "불만족 비율 (%)", "자살률", "-0.1765", "1.06e-01", "-0.1716", "1.16e-01"],
    ["", "만족도 점수", "자살률", "-0.0304", "7.82e-01", "0.0391", "7.23e-01"],
    ["여성", "만족 비율 (%)", "자살률", "-0.0422", "7.02e-01", "-0.0591", "5.91e-01"],
    ["", "불만족 비율 (%)", "자살률", "-0.0674", "5.40e-01", "-0.0450", "6.83e-01"],
    ["", "만족도 점수", "자살률", "0.0136", "9.01e-01", "-0.0180", "8.70e-01"]
]

pdf.set_font("Malgun", "", 8)
with pdf.table(col_widths=(15, 25, 25, 25, 25, 25, 25), text_align="CENTER") as table:
    row = table.row()
    pdf.set_font("Malgun", "B", 8)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(31, 78, 120)
    for h in table2_headers:
        row.cell(h)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Malgun", "", 7.5)
    for r_idx, r_data in enumerate(table2_rows):
        row = table.row()
        if r_idx % 2 == 1:
            pdf.set_fill_color(242, 242, 242)
        else:
            pdf.set_fill_color(255, 255, 255)
        for val in r_data:
            row.cell(val)

pdf.ln(4)
print_paragraph("또한, 연도별 시도별로 단년도 단위로 상관관계를 분석한 결과는 아래 표 3과 같습니다.")
pdf.ln(2)

# Table 3: Year-by-Year Correlation
table3_headers = ["연도", "독립 변수 (X)", "종속 변수 (Y)", "피어슨 r", "p-value", "스피어먼 ρ", "p-value"]
table3_rows = [
    ["2020년", "만족도 점수", "자살률", "-0.5036", "0.0393 *", "-0.4020", "1.10e-01"],
    ["", "만족 비율 (%)", "자살률", "-0.5226", "0.0314 *", "-0.3775", "1.35e-01"],
    ["2021년", "만족도 점수", "자살률", "0.0517", "8.44e-01", "0.2466", "3.40e-01"],
    ["", "만족 비율 (%)", "자살률", "-0.1126", "6.67e-01", "0.1683", "5.18e-01"],
    ["2022년", "만족도 점수", "자살률", "0.0785", "7.65e-01", "-0.1128", "6.66e-01"],
    ["", "만족 비율 (%)", "자살률", "-0.0077", "9.77e-01", "-0.1755", "5.01e-01"],
    ["2023년", "만족도 점수", "자살률", "-0.1120", "6.69e-01", "0.0454", "8.63e-01"],
    ["", "만족 비율 (%)", "자살률", "-0.2396", "3.54e-01", "-0.0331", "9.00e-01"],
    ["2024년", "만족도 점수", "자살률", "0.2299", "3.75e-01", "0.2232", "3.89e-01"],
    ["", "만족 비율 (%)", "자살률", "0.0386", "8.83e-01", "0.2074", "4.25e-01"]
]

pdf.set_font("Malgun", "", 8)
with pdf.table(col_widths=(15, 25, 25, 25, 25, 25, 25), text_align="CENTER") as table:
    row = table.row()
    pdf.set_font("Malgun", "B", 8)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(31, 78, 120)
    for h in table3_headers:
        row.cell(h)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Malgun", "", 7.5)
    for r_idx, r_data in enumerate(table3_rows):
        row = table.row()
        if r_idx % 2 == 1:
            pdf.set_fill_color(242, 242, 242)
        else:
            pdf.set_fill_color(255, 255, 255)
        for val in r_data:
            row.cell(val)

# New Page for Section 4 Images
pdf.add_page()

# Image 3
img_path3 = os.path.join(images_dir, "satisfaction_vs_suicide_scatter.png")
if os.path.exists(img_path3):
    pdf.image(img_path3, x=20, y=20, w=170)
    pdf.set_y(105)
    pdf.set_font("Malgun", "", 8.5)
    pdf.set_text_color(89, 89, 89)
    pdf.cell(0, 5, "<그림 3> 삶의 만족도 점수 vs 10만 명당 자살률 산점도 (성별/연도 통합)", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(10)

# Image 4
img_path4 = os.path.join(images_dir, "correlation_heatmap.png")
if os.path.exists(img_path4):
    pdf.image(img_path4, x=30, y=120, w=150)
    pdf.set_y(215)
    pdf.set_font("Malgun", "", 8.5)
    pdf.set_text_color(89, 89, 89)
    pdf.cell(0, 5, "<그림 4> 삶의 세부 만족도 지표 및 자살률 간의 상관관계 열지도(Heatmap)", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# New Page for Section 5
pdf.add_page()

# 5. 시도별 자살률과 만족도의 역설
print_heading("5. 시도별 자살률과 만족도의 역설 (The Satisfaction Paradox)")
print_paragraph("2024년 기준 17개 시도별 삶의 만족도 상세 통계 비율과 자살률 데이터를 대조하면 주관적 감정 만족도와 실제 사망 통계 간의 역설이 뚜렷하게 관찰됩니다.")
pdf.ln(2)

# Table 4: 2024 Regional Stats
regional_headers = ["순위", "행정구역", "만족도점수", "매우만족", "약간만족", "보통", "약간불만", "매우불만", "자살률"]
regional_rows = [
    ["1", "제주특별자치도", "3.44", "12.8%", "30.6%", "46.4%", "8.4%", "1.8%", "36.3"],
    ["2", "충청남도", "3.56", "16.3%", "30.7%", "46.7%", "4.9%", "1.3%", "34.8"],
    ["3", "전라남도", "3.34", "10.3%", "29.9%", "45.4%", "12.6%", "1.8%", "34.5"],
    ["4", "강원특별자치도", "3.42", "12.1%", "30.1%", "47.7%", "8.5%", "1.7%", "34.3"],
    ["5", "전북특별자치도", "3.38", "10.6%", "34.2%", "41.0%", "11.1%", "3.0%", "32.3"],
    ["6", "충청북도", "3.39", "13.0%", "27.7%", "47.0%", "10.0%", "2.2%", "31.8"],
    ["7", "경상북도", "3.35", "9.2%", "31.7%", "46.2%", "10.7%", "2.2%", "31.6"],
    ["8", "대전광역시", "3.52", "16.0%", "31.2%", "43.2%", "8.1%", "1.5%", "31.2"],
    ["8", "인천광역시", "3.21", "9.5%", "18.5%", "58.3%", "11.6%", "2.2%", "31.2"],
    ["10", "부산광역시", "3.32", "12.1%", "25.4%", "47.5%", "12.0%", "3.0%", "30.3"],
    ["11", "광주광역시", "3.41", "13.2%", "27.5%", "48.6%", "8.9%", "1.9%", "29.9"],
    ["12", "대구광역시", "3.21", "9.2%", "26.9%", "42.6%", "18.9%", "2.5%", "29.4"],
    ["13", "울산광역시", "3.43", "13.4%", "27.2%", "49.6%", "8.1%", "1.7%", "29.2"],
    ["14", "경상남도", "3.40", "10.4%", "32.5%", "45.2%", "10.3%", "1.5%", "28.5"],
    ["15", "경기도", "3.31", "10.3%", "25.4%", "51.4%", "10.7%", "2.3%", "28.2"],
    ["16", "서울특별시", "3.44", "12.4%", "33.9%", "41.8%", "9.3%", "2.6%", "24.1"],
    ["17", "세종특별자치시", "3.35", "16.5%", "26.3%", "36.4%", "17.4%", "3.4%", "23.0"]
]

pdf.set_font("Malgun", "", 7.5)
with pdf.table(col_widths=(10, 25, 20, 20, 20, 15, 20, 20, 20), text_align="CENTER") as table:
    row = table.row()
    pdf.set_font("Malgun", "B", 7.5)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(31, 78, 120)
    for h in regional_headers:
        row.cell(h)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Malgun", "", 7)
    for r_idx, r_data in enumerate(regional_rows):
        row = table.row()
        if r_idx % 2 == 1:
            pdf.set_fill_color(242, 242, 242)
        else:
            pdf.set_fill_color(255, 255, 255)
        for val in r_data:
            row.cell(val)

pdf.ln(5)

# New Page for Section 5 Images
pdf.add_page()

# Image 5
img_path5 = os.path.join(images_dir, "regional_comparison_bar.png")
if os.path.exists(img_path5):
    pdf.image(img_path5, x=20, y=20, w=170)
    pdf.set_y(105)
    pdf.set_font("Malgun", "", 8.5)
    pdf.set_text_color(89, 89, 89)
    pdf.cell(0, 5, "<그림 5> 2024년 시도별 자살률 및 삶의 만족도 비교 (자살률 순 정렬)", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(10)

# Image 6
img_path6 = os.path.join(images_dir, "satisfaction_vs_suicide_2024_annotated.png")
if os.path.exists(img_path6):
    pdf.image(img_path6, x=20, y=120, w=170)
    pdf.set_y(205)
    pdf.set_font("Malgun", "", 8.5)
    pdf.set_text_color(89, 89, 89)
    pdf.cell(0, 5, "<그림 6> 2024년 시도별 삶의 만족도 점수 vs 자살률 분포 (시도 산점도)", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

# New Page for Discussion & Conclusion
pdf.add_page()

pdf.ln(5)
print_paragraph("역설이 일어나는 요인은 세 가지로 분석됩니다.")
print_paragraph("1) 생태학적 오류 및 고령화 편향: 고령인구 비중이 높은 도(道) 지역은 노인 자살률이 높아 전체 평균 자살률이 올라가는 반면, 설문조사의 주관적 만족도는 일반 인구 분포를 따르므로 취약 노인층의 절망이 과소대표됩니다.")
print_paragraph("2) 사회적 비교 이론: 만족도가 높은 이웃들 사이에 홀로 고독이나 불만을 느끼는 개인은 그 소외감과 심리적 박탈감이 더욱 강하게 나타납니다.")
print_paragraph("3) 보통(무색무취) 비율과 자살률의 연계: 적극적으로 만족/불만족을 표현하지 않고 보통 비율이 높은 '정서적 무관심 지대'가 많은 지역일수록 실제 자살 예방 안전망이 부실하거나 취약 계층의 발굴이 어려워 자살 위험이 상승합니다.")
pdf.ln(5)

# 6. 결론 및 제언
print_heading("6. 결론 및 제언")
conclusions = [
    ("• 정신건강 정책 지표 다각화: ", "지역 복지 정책 수립 시 단순히 삶의 만족도 서베이 단독 지표만 활용하기보다는 자살 리스크 데이터와 노인 1인가구 등 실질 리스크 인자를 병합하여 위험 지역을 세밀하게 타겟팅해야 합니다."),
    ("• 무관심 소외 계층(보통 응답층) 발굴: ", "적극적 불만을 노출하지 않고 감정을 유보하거나 '보통'이라고 진술하여 통계적 사각지대에 놓이는 지역 소외 계층을 식별하기 위한 정서 발굴 프로그램 및 탐지망 도입이 필요합니다."),
    ("• 지방 및 도 지역 특화 복지망 및 성별 예방: ", "자살률이 매우 높게 지속되는 충남, 제주, 전남, 강원 등 도 지역의 취약 가구 밀착 복지 강화와 함께, 남성 만족도 대비 자살률이 2.5배 가량 높은 성별 불균형 문제를 해소하기 위한 남성 특화 프로그램(경제적 실직, 남성 고립 정서 대상) 확대가 강력히 요구됩니다.")
]
for prefix, body in conclusions:
    print_paragraph(body, bold_prefix=prefix)

# Output PDF
pdf.output(pdf_path)
print(f"PDF document saved successfully to {pdf_path}!")
