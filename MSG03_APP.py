# ============================================================
#  MSG-03  AI 기반 품질 검사 자동화
#  경영시스템(주) | AI 담당 최병대 | 010-3533-9030
#  AI program base 300ea | Ver 1.0 | 2026.03
#  실행: streamlit run MSG03_APP.py
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image, ImageDraw, ImageFilter
import os, io, random
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ──────────────────────────────────────────────
# 페이지 설정
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="MSG-03 | AI 품질 검사 자동화",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "25_output_data")
INPUT_DIR  = os.path.join(BASE_DIR, "20_input_data")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(INPUT_DIR,  exist_ok=True)

# ──────────────────────────────────────────────
# 자동 저장 공통 함수 (10,000개 프로그램 공통)
# ──────────────────────────────────────────────
def auto_save_excel(df, prefix="결과"):
    try:
        ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"MSG03_{prefix}_{ts}.xlsx"
        fpath = os.path.join(OUTPUT_DIR, fname)
        wb    = openpyxl.Workbook()
        ws    = wb.active
        ws.title = prefix[:30]
        navy  = PatternFill(start_color="0D1B4B", end_color="0D1B4B", fill_type="solid")
        light = PatternFill(start_color="EEF2F7", end_color="EEF2F7", fill_type="solid")
        white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        thin  = Border(left=Side(style='thin'), right=Side(style='thin'),
                       top=Side(style='thin'),  bottom=Side(style='thin'))
        # 제목
        ws.merge_cells(f"A1:{chr(64+max(len(df.columns),1))}1")
        ws["A1"] = f"MSG-03 AI 품질검사 | {prefix} | {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        ws["A1"].font      = Font(name='맑은 고딕', bold=True, color='FFD700', size=11)
        ws["A1"].fill      = navy
        ws["A1"].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 28
        # 헤더
        for ci, col in enumerate(df.columns, 1):
            c = ws.cell(row=2, column=ci, value=col)
            c.font      = Font(name='맑은 고딕', bold=True, color='FFFFFF', size=10)
            c.fill      = navy
            c.border    = thin
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[chr(64+ci)].width = max(15, len(str(col))+4)
        ws.row_dimensions[2].height = 22
        # 데이터
        for ri, row in enumerate(df.itertuples(index=False), 3):
            fill = light if ri % 2 == 0 else white
            for ci, val in enumerate(row, 1):
                c = ws.cell(row=ri, column=ci, value=str(val) if val is not None else '')
                c.font      = Font(name='맑은 고딕', size=10)
                c.fill      = fill
                c.border    = thin
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws.row_dimensions[ri].height = 20
        wb.save(fpath)
        return fpath, fname
    except Exception as e:
        return None, str(e)

# ──────────────────────────────────────────────
# CSS (MSG 공통 디자인)
# ──────────────────────────────────────────────
st.markdown("""
<style>
  body, .stApp { background-color:#EEF2F7; font-family:'Malgun Gothic',sans-serif; }

  .top-banner {
    background:linear-gradient(135deg,#0D1B4B 0%,#1A3080 60%,#1565C0 100%);
    border-radius:12px; padding:18px 28px; margin-bottom:18px;
    box-shadow:0 4px 16px rgba(0,0,0,0.25);
  }
  .top-banner h1 { color:#FFD700; font-size:22px; font-weight:800; margin:0; }
  .top-banner p  { color:#90CAF9; font-size:12px; margin:4px 0 0 0; }

  .kpi-card {
    background:white; border-radius:10px; padding:16px 20px;
    box-shadow:0 2px 10px rgba(0,0,0,0.08); border-top:4px solid #1A3080;
    text-align:center;
  }
  .kpi-title { color:#666; font-size:12px; margin-bottom:6px; }
  .kpi-value { color:#0D1B4B; font-size:28px; font-weight:800; }
  .kpi-unit  { color:#888; font-size:12px; }
  .kpi-sub   { color:#1565C0; font-size:11px; margin-top:4px; }

  .sec-title {
    background:#0D1B4B; color:#FFD700;
    padding:8px 16px; border-radius:6px;
    font-size:15px; font-weight:700; margin-bottom:14px;
  }

  .grade-critical { background:#FFEBEE; border:2px solid #C62828; border-radius:8px; padding:12px; text-align:center; }
  .grade-major    { background:#FFF3E0; border:2px solid #E65100; border-radius:8px; padding:12px; text-align:center; }
  .grade-minor    { background:#FFFDE7; border:2px solid #F9A825; border-radius:8px; padding:12px; text-align:center; }
  .grade-pass     { background:#E8F5E9; border:2px solid #2E7D32; border-radius:8px; padding:12px; text-align:center; }
  .grade-hold     { background:#F3E5F5; border:2px solid #6A1B9A; border-radius:8px; padding:12px; text-align:center; }

  .info-box {
    background:#E3F2FD; border-left:4px solid #1565C0;
    border-radius:6px; padding:12px 16px; margin:10px 0; font-size:13px; color:#0D1B4B;
  }
  .warn-box {
    background:#FFF8E1; border-left:4px solid #FFD700;
    border-radius:6px; padding:12px 16px; margin:10px 0; font-size:13px; color:#5D4037;
  }
  .footer-bar {
    background:#0D1B4B; color:#90CAF9;
    padding:10px 20px; border-radius:8px;
    font-size:11px; text-align:center; margin-top:30px;
  }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 샘플 데이터 자동 생성 (앱 내부 - 별도 파일 불필요)
# ──────────────────────────────────────────────
@st.cache_data
def generate_inspection_data():
    random.seed(42); np.random.seed(42)
    products  = ['엔진 브라켓','도어 패널','변속기 케이스','브레이크 디스크','서스펜션 암',
                 'PCB 기판','커넥터','하우징','샤프트','기어']
    lines     = ['1호 라인','2호 라인','3호 라인','4호 라인','5호 라인']
    inspectors= ['김철수','이영희','박민준','최지원','정수빈']
    defects   = ['스크래치','크랙','치수불량','변색','이물질','기포','찍힘','뒤틀림']
    grades    = ['양호','양호','양호','양호','Minor','Minor','Major','Critical','보류']
    base = datetime(2024, 1, 1)
    rows = []
    for i in range(500):
        grade   = random.choice(grades)
        result  = '합격' if grade in ['양호','Minor'] else ('불합격' if grade in ['Critical','Major'] else '보류')
        defect  = random.choice(defects) if grade != '양호' else '없음'
        product = random.choice(products)
        rows.append({
            '검사번호':     f'QI-2024-{i+1:05d}',
            '검사일시':     base + timedelta(days=random.randint(0,730), hours=random.randint(8,17)),
            '제품명':       product,
            '라인':         random.choice(lines),
            '검사자':       random.choice(inspectors),
            '불량등급':     grade,
            '판정결과':     result,
            '불량유형':     defect,
            '외관점수':     round(random.uniform(60,100),1),
            '치수편차(mm)': round(random.uniform(0,0.5),3),
            '중량편차(g)':  round(random.uniform(0,5),2),
            '검사시간(초)': random.randint(10,120),
            '처리상태':     random.choice(['완료','보류중','재검사']) if result != '합격' else '완료',
        })
    df = pd.DataFrame(rows)
    df['검사일시'] = pd.to_datetime(df['검사일시'])
    return df.sort_values('검사일시').reset_index(drop=True)

@st.cache_data
def generate_dimension_data():
    """치수 검사 샘플 데이터"""
    random.seed(10); np.random.seed(10)
    items = ['전장(mm)','전폭(mm)','전고(mm)','직경(mm)','두께(mm)','홀간격(mm)']
    rows  = []
    for i in range(100):
        for item in items:
            nominal = {'전장(mm)':100.0,'전폭(mm)':50.0,'전고(mm)':30.0,
                       '직경(mm)':25.0,'두께(mm)':5.0,'홀간격(mm)':20.0}[item]
            tol     = 0.1
            measured= round(nominal + random.uniform(-0.15, 0.15), 3)
            dev     = round(measured - nominal, 3)
            ok      = '합격' if abs(dev) <= tol else '불합격'
            rows.append({
                '측정번호': f'DIM-{i+1:04d}',
                '측정항목': item,
                '기준값':   nominal,
                '공차(±)':  tol,
                '측정값':   measured,
                '편차':     dev,
                '판정':     ok,
            })
    return pd.DataFrame(rows)

def create_inspection_image(grade='양호'):
    """검사 이미지 시뮬레이션"""
    img  = Image.new('RGB', (400, 300), color=(220, 220, 220))
    draw = ImageDraw.Draw(img)
    # 제품 외형
    draw.rectangle([50, 50, 350, 250], outline=(100,100,100), width=3, fill=(200,200,200))
    draw.rectangle([80, 80, 320, 220], outline=(150,150,150), width=1, fill=(210,210,210))
    # 결함 표시
    if grade == 'Critical':
        draw.line([(150,100),(250,200)], fill=(255,0,0), width=4)
        draw.line([(250,100),(150,200)], fill=(255,0,0), width=4)
    elif grade == 'Major':
        for _ in range(3):
            x = random.randint(100,300); y = random.randint(80,220)
            draw.ellipse([x-8,y-8,x+8,y+8], fill=(255,100,0))
    elif grade == 'Minor':
        x = random.randint(120,280); y = random.randint(90,210)
        draw.line([(x,y),(x+30,y+5)], fill=(200,150,0), width=2)
    elif grade == '보류':
        draw.rectangle([120,100,280,200], outline=(128,0,128), width=3)
    img = img.filter(ImageFilter.GaussianBlur(radius=0.5))
    return img

# ──────────────────────────────────────────────
# 데이터 로드
# ──────────────────────────────────────────────
df     = generate_inspection_data()
dim_df = generate_dimension_data()

# ──────────────────────────────────────────────
# 사이드바
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ✅ MSG-03 메뉴")
    st.markdown("---")
    menu = st.radio("", [
        "📊 검사 현황 대시보드",
        "📷 이미지 검사 자동화",
        "📏 치수 검사 자동화",
        "⚖️ 중량/기능 검사",
        "✅ 검사 기준 관리",
        "📋 검사 성적서 발행",
        "📈 불량 트렌드 분석",
        "🤖 AI 판정 & 인사이트",
    ], label_visibility="hidden")
    st.markdown("---")
    st.markdown("### 📂 샘플 데이터")
    st.caption(f"총 {len(df):,}건 자동생성")
    st.markdown("---")
    st.markdown("""
    <div style='font-size:11px;color:#aaa;'>
    경영시스템(주)<br>AI 담당 최병대<br>010-3533-9030
    </div>""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 공통 헤더
# ──────────────────────────────────────────────
st.markdown("""
<div class="top-banner">
  <h1>✅ MSG-03  AI 기반 품질 검사 자동화 | AI Quality Inspection Automation</h1>
  <p>경영시스템(주) | AI 담당 최병대 | 010-3533-9030 | AI program base 300ea &nbsp;|&nbsp; Ver 1.0 &nbsp;|&nbsp; 2026.03</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════
# ① 검사 현황 대시보드
# ══════════════════════════════════════════════
if menu == "📊 검사 현황 대시보드":
    st.markdown('<div class="sec-title">📊 품질 검사 현황 대시보드 | Quality Inspection Dashboard</div>', unsafe_allow_html=True)

    total     = len(df)
    pass_cnt  = len(df[df['판정결과']=='합격'])
    fail_cnt  = len(df[df['판정결과']=='불합격'])
    hold_cnt  = len(df[df['판정결과']=='보류'])
    pass_rate = round(pass_cnt/total*100, 1)
    critical  = len(df[df['불량등급']=='Critical'])
    avg_time  = round(df['검사시간(초)'].mean(), 1)

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    kpis = [
        (c1, "총 검사건수 / Total",    f"{total:,}",     "건",  ""),
        (c2, "합격건수 / Pass",         f"{pass_cnt:,}",  "건",  f"▲ {pass_rate}%"),
        (c3, "불합격건수 / Fail",       f"{fail_cnt:,}",  "건",  f"▼ {round(fail_cnt/total*100,1)}%"),
        (c4, "보류건수 / Hold",         f"{hold_cnt:,}",  "건",  "재검사 필요"),
        (c5, "Critical 불량",          f"{critical:,}",  "건",  "즉시 조치 필요"),
        (c6, "평균 검사시간",           f"{avg_time}",    "초",  "▼ 목표: 30초"),
    ]
    for col, title, val, unit, sub in kpis:
        col.markdown(f"""
        <div class="kpi-card">
          <div class="kpi-title">{title}</div>
          <div class="kpi-value">{val}<span class="kpi-unit"> {unit}</span></div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        monthly = df.copy()
        monthly['월'] = monthly['검사일시'].dt.to_period('M').astype(str)
        mg = monthly.groupby(['월','판정결과']).size().reset_index(name='건수')
        fig = px.bar(mg, x='월', y='건수', color='판정결과',
                     color_discrete_map={'합격':'#1A3080','불합격':'#F44336','보류':'#FF9800'},
                     title='월별 검사 현황 | Monthly Trend', barmode='stack')
        fig.update_layout(plot_bgcolor='white', height=300)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        grade_cnt = df['불량등급'].value_counts().reset_index()
        grade_cnt.columns = ['등급','건수']
        color_map = {'양호':'#1A3080','Minor':'#FF9800','Major':'#F44336',
                     'Critical':'#B71C1C','보류':'#7B1FA2'}
        fig2 = px.pie(grade_cnt, names='등급', values='건수',
                      title='불량 등급 분포 | Defect Grade Distribution',
                      color='등급', color_discrete_map=color_map)
        fig2.update_layout(height=300)
        st.plotly_chart(fig2, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        line_grp = df.groupby('라인')['판정결과'].apply(
            lambda x: round((x=='불합격').sum()/len(x)*100,1)).reset_index()
        line_grp.columns = ['라인','불합격률(%)']
        fig3 = px.bar(line_grp, x='라인', y='불합격률(%)',
                      title='라인별 불합격률 | Fail Rate by Line',
                      color='불합격률(%)', color_continuous_scale='Reds')
        fig3.update_layout(plot_bgcolor='white', height=280)
        st.plotly_chart(fig3, use_container_width=True)

    with c4:
        defect_grp = df[df['불량유형']!='없음']['불량유형'].value_counts().reset_index()
        defect_grp.columns = ['불량유형','건수']
        fig4 = px.bar(defect_grp, x='건수', y='불량유형', orientation='h',
                      title='불량유형 순위 | Defect Type Ranking',
                      color='건수', color_continuous_scale='Blues_r')
        fig4.update_layout(plot_bgcolor='white', height=280)
        st.plotly_chart(fig4, use_container_width=True)

    # 자동저장
    kpi_df = pd.DataFrame({
        '항목':['총검사','합격','불합격','보류','합격률(%)','Critical','평균검사시간(초)'],
        '값':  [total, pass_cnt, fail_cnt, hold_cnt, pass_rate, critical, avg_time]
    })
    fpath, fname = auto_save_excel(kpi_df, '대시보드_KPI')
    if fpath:
        st.success(f'✅ 자동저장: 25_output_data/{fname}')

# ══════════════════════════════════════════════
# ② 이미지 검사 자동화
# ══════════════════════════════════════════════
elif menu == "📷 이미지 검사 자동화":
    st.markdown('<div class="sec-title">📷 이미지 검사 자동화 | Image Inspection Automation</div>', unsafe_allow_html=True)

    st.markdown('<div class="info-box">📌 AI 비전 카메라로 제품 외관을 자동 검사합니다. 이미지를 업로드하거나 샘플을 선택하세요. | AI vision camera automatically inspects product appearance.</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("#### 🖼️ 검사 이미지 | Inspection Image")
        uploaded = st.file_uploader("이미지 업로드 (JPG/PNG)", type=['jpg','jpeg','png','bmp'])
        use_sample = st.checkbox("✅ 샘플 이미지 사용", value=True)
        sample_grade = st.selectbox("샘플 불량등급 선택",
                                    ['양호','Minor','Major','Critical','보류'])
        product_name = st.text_input("제품명 | Product", value="엔진 브라켓")
        lot_no       = st.text_input("LOT No.", value=f"LOT-{datetime.now().strftime('%Y%m%d')}-001")

        if use_sample or uploaded is None:
            img = create_inspection_image(grade=sample_grade)
            st.image(img, caption=f"샘플 이미지 ({sample_grade})", use_column_width=True)
        else:
            img = Image.open(uploaded)
            st.image(img, caption="업로드된 이미지", use_column_width=True)

    with col2:
        st.markdown("#### 🤖 AI 분석 결과 | AI Analysis Result")
        std_option = st.selectbox("검사 기준 | Standard",
                                  ['사내기준 IN-SPEC','KS B ISO 1302','고객사 기준','IPC-A-610'])
        sensitivity = st.slider("AI 감도 | Sensitivity", 1, 10, 7)

        if st.button("🔍 AI 검사 실행 | Run AI Inspection", type="primary"):
            import time; time.sleep(1.0)

            grade    = sample_grade if use_sample else random.choice(['양호','Minor','Major'])
            conf     = round(random.uniform(88, 99), 1)
            score    = round(random.uniform(60, 100) if grade != '양호' else random.uniform(85, 100), 1)
            defect_t = random.choice(['스크래치','크랙','변색','이물질']) if grade != '양호' else '없음'
            result   = '합격' if grade in ['양호','Minor'] else ('불합격' if grade == 'Critical' else '보류')

            # 등급별 표시
            css_map = {'양호':'grade-pass','Minor':'grade-minor',
                       'Major':'grade-major','Critical':'grade-critical','보류':'grade-hold'}
            label_map = {'양호':'✅ 양호 GOOD','Minor':'⚠️ Minor 경미',
                         'Major':'🔶 Major 중결함','Critical':'🔴 Critical 치명','보류':'🟣 보류 HOLD'}

            st.markdown(f"""
            <div class="{css_map[grade]}">
              <div style="font-size:22px; font-weight:800;">{label_map[grade]}</div>
              <div style="margin-top:8px; font-size:13px;">
              판정결과: <b>{result}</b><br>
              외관점수: <b>{score}점</b><br>
              불량유형: <b>{defect_t}</b><br>
              AI 신뢰도: <b>{conf}%</b><br>
              적용기준: <b>{std_option}</b>
              </div>
            </div>""", unsafe_allow_html=True)

            if grade != '양호':
                st.markdown(f"""
                <div class="warn-box">
                ⚠️ <b>조치 권고사항 | Action Required</b><br>
                - 불량 등급: {grade}<br>
                - 불량 유형: {defect_t}<br>
                - {'즉시 라인 정지 및 전수검사 실시' if grade=='Critical' else '해당 LOT 분리 및 재검사 실시'}<br>
                - 성적서 발행 후 품질팀 보고
                </div>""", unsafe_allow_html=True)

            # 결함 위치 맵
            st.markdown("#### 📍 결함 위치 맵 | Defect Location Map")
            fig = go.Figure()
            fig.add_shape(type="rect", x0=0, y0=0, x1=400, y1=300,
                         fillcolor="#DDEEFF", opacity=0.5, line_color="#1A3080", line_width=2)
            fig.add_annotation(x=200, y=150, text="제품 | Product",
                              font=dict(size=14, color="#1A3080"), showarrow=False)
            if grade != '양호':
                for _ in range(random.randint(1,3)):
                    fx = random.randint(80,320); fy = random.randint(60,240)
                    sz = random.randint(15,35)
                    fig.add_shape(type="rect", x0=fx-sz, y0=fy-sz, x1=fx+sz, y1=fy+sz,
                                 fillcolor="red", opacity=0.3, line_color="red", line_width=2)
                    fig.add_annotation(x=fx, y=fy-sz-10,
                                      text=f"{defect_t}({grade})",
                                      font=dict(size=10, color="red"), showarrow=True,
                                      arrowhead=2, arrowcolor="red")
            fig.update_layout(width=400, height=250, plot_bgcolor='white',
                            xaxis=dict(showticklabels=False, range=[0,400]),
                            yaxis=dict(showticklabels=False, range=[0,300]),
                            margin=dict(l=0,r=0,t=10,b=0))
            st.plotly_chart(fig, use_container_width=True)

            # 자동저장
            result_df = pd.DataFrame({
                '항목':['제품명','LOT No.','불량등급','판정결과','외관점수','불량유형','AI신뢰도(%)','검사기준','검사일시'],
                '내용':[product_name, lot_no, grade, result, score, defect_t, conf, std_option,
                       datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
            })
            fpath, fname = auto_save_excel(result_df, '이미지검사결과')
            if fpath:
                st.success(f'✅ 자동저장: 25_output_data/{fname}')

# ══════════════════════════════════════════════
# ③ 치수 검사 자동화
# ══════════════════════════════════════════════
elif menu == "📏 치수 검사 자동화":
    st.markdown('<div class="sec-title">📏 치수 검사 자동화 | Dimension Inspection Automation</div>', unsafe_allow_html=True)

    st.markdown('<div class="info-box">📌 제품 치수를 측정하고 공차 범위 내 합격 여부를 자동 판정합니다. | Automatically judges dimensional conformance within tolerance.</div>', unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📝 직접 입력 | Manual Input", "📂 Excel 업로드 | Excel Upload"])

    with tab1:
        st.markdown("#### 📐 치수 직접 입력 | Manual Dimension Input")
        product_d = st.text_input("제품명 | Product", value="브레이크 디스크")
        lot_d     = st.text_input("LOT No.", value=f"LOT-{datetime.now().strftime('%Y%m%d')}-001")

        items = ['전장(mm)','전폭(mm)','전고(mm)','직경(mm)','두께(mm)','홀간격(mm)']
        nominals = [100.0, 50.0, 30.0, 25.0, 5.0, 20.0]
        tolerances = [0.1, 0.1, 0.05, 0.05, 0.02, 0.05]

        rows = []
        cols = st.columns(3)
        for i, (item, nom, tol) in enumerate(zip(items, nominals, tolerances)):
            with cols[i % 3]:
                measured = st.number_input(
                    f"{item} (기준:{nom}±{tol})",
                    min_value=float(nom-1), max_value=float(nom+1),
                    value=float(nom + round(random.uniform(-0.08, 0.08), 3)),
                    format="%.3f", key=f"dim_{i}"
                )
                dev    = round(measured - nom, 3)
                ok     = '✅합격' if abs(dev) <= tol else '❌불합격'
                rows.append({'항목':item, '기준값':nom, '공차(±)':tol,
                             '측정값':measured, '편차':dev, '판정':ok})

        if st.button("📏 치수 판정 실행 | Run Dimension Check", type="primary"):
            result_df2 = pd.DataFrame(rows)
            pass_all   = all('합격' in r for r in result_df2['판정'])
            fail_items = result_df2[result_df2['판정'].str.contains('불합격')]

            if pass_all:
                st.markdown("""
                <div class="grade-pass">
                  <div style="font-size:22px;font-weight:800;">✅ 전 항목 합격 | ALL PASS</div>
                </div>""", unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class="grade-critical">
                  <div style="font-size:22px;font-weight:800;">❌ 불합격 | FAIL</div>
                  <div style="font-size:13px;margin-top:6px;">불합격 항목: {', '.join(fail_items['항목'].tolist())}</div>
                </div>""", unsafe_allow_html=True)

            st.markdown("#### 📊 치수 측정 결과 | Measurement Results")
            st.dataframe(result_df2, use_container_width=True, hide_index=True)

            # 편차 차트
            fig = px.bar(result_df2, x='항목', y='편차',
                        title='치수 편차 현황 | Dimension Deviation',
                        color='판정', color_discrete_map={'✅합격':'#1A3080','❌불합격':'#F44336'})
            fig.add_hline(y=0.1,  line_dash="dash", line_color="orange", annotation_text="상한 공차")
            fig.add_hline(y=-0.1, line_dash="dash", line_color="orange", annotation_text="하한 공차")
            fig.update_layout(plot_bgcolor='white', height=280)
            st.plotly_chart(fig, use_container_width=True)

            # 자동저장
            save_df2 = result_df2.copy()
            save_df2.insert(0, '제품명', product_d)
            save_df2.insert(1, 'LOT No.', lot_d)
            fpath, fname = auto_save_excel(save_df2, '치수검사결과')
            if fpath:
                st.success(f'✅ 자동저장: 25_output_data/{fname}')

    with tab2:
        st.markdown("#### 📂 Excel 파일 업로드 | Excel File Upload")
        st.markdown('<div class="info-box">📌 Excel 파일 형식: 항목, 기준값, 공차, 측정값 열이 필요합니다.</div>', unsafe_allow_html=True)
        uploaded_excel = st.file_uploader("Excel 파일 업로드", type=['xlsx','xls'])

        if uploaded_excel:
            df_excel = pd.read_excel(uploaded_excel)
            st.dataframe(df_excel, use_container_width=True)
            fpath, fname = auto_save_excel(df_excel, '치수검사_업로드')
            if fpath:
                st.success(f'✅ 자동저장: 25_output_data/{fname}')
        else:
            st.markdown("#### 📋 샘플 데이터 미리보기 | Sample Data Preview")
            st.dataframe(dim_df.head(20), use_container_width=True, hide_index=True)

            # 샘플 자동저장
            fpath, fname = auto_save_excel(dim_df, '치수검사_샘플')
            if fpath:
                st.success(f'✅ 자동저장: 25_output_data/{fname}')

# ══════════════════════════════════════════════
# ④ 중량/기능 검사
# ══════════════════════════════════════════════
elif menu == "⚖️ 중량/기능 검사":
    st.markdown('<div class="sec-title">⚖️ 중량/기능 검사 | Weight & Function Inspection</div>', unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["⚖️ 중량 검사 | Weight", "🔧 기능 검사 | Function"])

    with tab1:
        st.markdown("#### ⚖️ 중량 검사 | Weight Inspection")
        col1, col2 = st.columns(2)
        with col1:
            product_w  = st.text_input("제품명", value="엔진 브라켓", key="wt_prod")
            nominal_w  = st.number_input("기준 중량 (g)", value=500.0, format="%.2f")
            tolerance_w= st.number_input("공차 ± (g)", value=5.0, format="%.2f")
            sample_cnt = st.number_input("측정 샘플 수", min_value=1, max_value=100, value=10)

        with col2:
            if st.button("⚖️ 중량 검사 실행", type="primary"):
                weights = [round(nominal_w + random.uniform(-8, 8), 2) for _ in range(int(sample_cnt))]
                w_df = pd.DataFrame({
                    '샘플번호': [f'S-{i+1:03d}' for i in range(int(sample_cnt))],
                    '측정중량(g)': weights,
                    '편차(g)': [round(w - nominal_w, 2) for w in weights],
                    '판정': ['✅합격' if abs(w-nominal_w) <= tolerance_w else '❌불합격' for w in weights]
                })
                pass_w = len(w_df[w_df['판정'].str.contains('합격')])
                st.metric("합격률", f"{round(pass_w/len(w_df)*100,1)}%")
                st.dataframe(w_df, use_container_width=True, hide_index=True)

                fig = px.scatter(w_df, x='샘플번호', y='측정중량(g)',
                                color='판정', color_discrete_map={'✅합격':'#1A3080','❌불합격':'#F44336'},
                                title='중량 측정 결과')
                fig.add_hline(y=nominal_w+tolerance_w, line_dash="dash", line_color="orange")
                fig.add_hline(y=nominal_w-tolerance_w, line_dash="dash", line_color="orange")
                fig.add_hline(y=nominal_w, line_color="green")
                fig.update_layout(plot_bgcolor='white', height=280)
                st.plotly_chart(fig, use_container_width=True)

                fpath, fname = auto_save_excel(w_df, '중량검사결과')
                if fpath:
                    st.success(f'✅ 자동저장: 25_output_data/{fname}')

    with tab2:
        st.markdown("#### 🔧 기능 검사 항목 | Function Test Items")
        func_items = {
            '내압 테스트 (MPa)':   (10.0, 0.5),
            '인장강도 (MPa)':      (350.0, 20.0),
            '경도 (HRC)':          (45.0, 3.0),
            '표면조도 Ra (μm)':    (1.6, 0.2),
            '토크 (N·m)':          (50.0, 5.0),
        }
        func_rows = []
        for item, (nom, tol) in func_items.items():
            measured = round(nom + random.uniform(-tol*1.5, tol*1.5), 2)
            ok = '✅합격' if abs(measured-nom) <= tol else '❌불합격'
            func_rows.append({'검사항목':item,'기준값':nom,'공차(±)':tol,
                              '측정값':measured,'편차':round(measured-nom,2),'판정':ok})
        func_df = pd.DataFrame(func_rows)

        if st.button("🔧 기능 검사 실행", type="primary"):
            st.dataframe(func_df, use_container_width=True, hide_index=True)
            pass_all_f = all('합격' in r for r in func_df['판정'])
            if pass_all_f:
                st.success("✅ 전 항목 합격 | ALL FUNCTION TEST PASS")
            else:
                st.error("❌ 불합격 항목 있음 | FUNCTION TEST FAIL")

            fpath, fname = auto_save_excel(func_df, '기능검사결과')
            if fpath:
                st.success(f'✅ 자동저장: 25_output_data/{fname}')

# ══════════════════════════════════════════════
# ⑤ 검사 기준 관리
# ══════════════════════════════════════════════
elif menu == "✅ 검사 기준 관리":
    st.markdown('<div class="sec-title">✅ 검사 기준 관리 | Inspection Standard Management</div>', unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📋 기준 조회","➕ 기준 등록","📤 기준 내보내기"])

    with tab1:
        standards_data = pd.DataFrame({
            '기준코드':  ['STD-001','STD-002','STD-003','STD-004','STD-005'],
            '제품군':    ['금속부품','플라스틱','전자부품','고무부품','복합소재'],
            '적용규격':  ['KS B ISO','사내기준','IPC-A-610','KS M','ASTM'],
            '외관기준':  ['스크래치 0.1mm이하','변색없음','납볼 직경±0.05mm','기포없음','균열없음'],
            '치수공차':  ['±0.1mm','±0.2mm','±0.05mm','±0.5mm','±0.15mm'],
            '중량공차':  ['±1%','±2%','±0.5%','±3%','±1.5%'],
            '합격기준':  ['95%이상','90%이상','98%이상','92%이상','95%이상'],
            '개정일':    ['2024-01-15','2024-03-20','2024-02-10','2023-12-01','2024-04-05'],
        })
        st.dataframe(standards_data, use_container_width=True, hide_index=True)
        fpath, fname = auto_save_excel(standards_data, '검사기준_조회')
        if fpath:
            st.success(f'✅ 자동저장: 25_output_data/{fname}')

    with tab2:
        st.markdown("#### ➕ 새 검사 기준 등록 | Register New Standard")
        c1, c2 = st.columns(2)
        with c1:
            new_code    = st.text_input("기준코드", value="STD-006")
            new_product = st.text_input("제품군", value="신규 부품")
            new_std     = st.selectbox("적용규격", ['KS B ISO','사내기준','IPC-A-610','ASTM','고객사기준'])
        with c2:
            new_appearance = st.text_input("외관기준", value="스크래치 없음")
            new_dim  = st.text_input("치수공차", value="±0.1mm")
            new_weight = st.text_input("중량공차", value="±1%")
            new_pass = st.text_input("합격기준", value="95%이상")

        if st.button("💾 기준 저장 | Save Standard", type="primary"):
            new_row = pd.DataFrame({
                '기준코드':[new_code],'제품군':[new_product],'적용규격':[new_std],
                '외관기준':[new_appearance],'치수공차':[new_dim],
                '중량공차':[new_weight],'합격기준':[new_pass],
                '개정일':[str(datetime.today().date())]
            })
            fpath, fname = auto_save_excel(new_row, '검사기준_신규등록')
            if fpath:
                st.success(f'✅ 저장 완료: 25_output_data/{fname}')
            st.dataframe(new_row, use_container_width=True, hide_index=True)

    with tab3:
        st.markdown("#### 📤 검사 기준 내보내기 | Export Standards")
        buf = io.BytesIO()
        standards_data = pd.DataFrame({
            '기준코드':  ['STD-001','STD-002','STD-003','STD-004','STD-005'],
            '제품군':    ['금속부품','플라스틱','전자부품','고무부품','복합소재'],
            '적용규격':  ['KS B ISO','사내기준','IPC-A-610','KS M','ASTM'],
        })
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            standards_data.to_excel(writer, index=False, sheet_name='검사기준')
        st.download_button("📥 Excel 다운로드", data=buf.getvalue(),
                          file_name=f"검사기준_{datetime.now().strftime('%Y%m%d')}.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════
# ⑥ 검사 성적서 발행
# ══════════════════════════════════════════════
elif menu == "📋 검사 성적서 발행":
    st.markdown('<div class="sec-title">📋 검사 성적서 발행 | Inspection Certificate</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 📝 기본 정보 | Basic Info")
        cert_no   = st.text_input("성적서번호", value=f"QI-CERT-{datetime.now().strftime('%Y%m%d')}-001")
        insp_date = st.date_input("검사일", value=datetime.today())
        client    = st.text_input("고객사 | Client", value="현대자동차(주)")
        product   = st.text_input("제품명 | Product", value="엔진 브라켓")
        lot_no    = st.text_input("LOT No.", value="LOT-20240115-001")
        qty       = st.number_input("검사수량 | Qty", value=100)
        pass_qty  = st.number_input("합격수량 | Pass Qty", value=97)

    with col2:
        st.markdown("#### 👤 검사자 정보 | Inspector Info")
        inspector = st.text_input("검사자 | Inspector", value="김철수")
        dept      = st.text_input("부서 | Dept", value="품질관리팀")
        company   = st.text_input("회사 | Company", value="경영시스템(주)")
        std_cert  = st.selectbox("적용규격", ['KS B ISO 1302','사내기준','고객사 기준','IPC-A-610'])
        insp_type = st.multiselect("검사항목", ['외관','치수','중량','기능'], default=['외관','치수'])
        result_c  = st.selectbox("최종판정", ['합격 PASS','불합격 FAIL','조건부합격','보류 HOLD'])
        remarks   = st.text_area("특기사항", value="검사 결과 이상 없음.", height=60)

    if st.button("📋 성적서 생성 및 저장", type="primary"):
        cert_df = pd.DataFrame({
            '항목': ['성적서번호','검사일','고객사','제품명','LOT No.','검사수량',
                    '합격수량','불합격수량','합격률(%)','검사자','부서','회사',
                    '적용규격','검사항목','최종판정','특기사항'],
            '내용': [cert_no, str(insp_date), client, product, lot_no, qty,
                    pass_qty, qty-pass_qty, round(pass_qty/qty*100,1),
                    inspector, dept, company, std_cert,
                    '+'.join(insp_type), result_c, remarks]
        })
        st.dataframe(cert_df, use_container_width=True, hide_index=True)
        fpath, fname = auto_save_excel(cert_df, '검사성적서')
        if fpath:
            st.success(f'✅ 성적서 저장 완료: 25_output_data/{fname}')

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            cert_df.to_excel(writer, index=False, sheet_name='성적서')
        st.download_button("📥 성적서 다운로드", data=buf.getvalue(),
                          file_name=fname,
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════
# ⑦ 불량 트렌드 분석
# ══════════════════════════════════════════════
elif menu == "📈 불량 트렌드 분석":
    st.markdown('<div class="sec-title">📈 불량 트렌드 분석 | Defect Trend Analysis</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        df_trend = df.copy()
        df_trend['월'] = df_trend['검사일시'].dt.to_period('M').astype(str)
        trend = df_trend.groupby('월').apply(
            lambda x: round((x['판정결과']=='불합격').sum()/len(x)*100,1)).reset_index()
        trend.columns = ['월','불합격률(%)']
        fig = px.line(trend, x='월', y='불합격률(%)',
                     title='월별 불합격률 추이 | Monthly Fail Rate Trend',
                     markers=True, line_shape='spline')
        fig.add_hline(y=5, line_dash="dash", line_color="red", annotation_text="목표 5%")
        fig.update_traces(line_color='#1A3080', marker_color='#FFD700', marker_size=8)
        fig.update_layout(plot_bgcolor='white', height=300)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        defect_monthly = df_trend[df_trend['불량유형']!='없음'].groupby(
            ['월','불량유형']).size().reset_index(name='건수')
        fig2 = px.bar(defect_monthly, x='월', y='건수', color='불량유형',
                     title='월별 불량유형 분포 | Defect Type by Month',
                     barmode='stack')
        fig2.update_layout(plot_bgcolor='white', height=300)
        st.plotly_chart(fig2, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        prod_fail = df.groupby('제품명').apply(
            lambda x: round((x['판정결과']=='불합격').sum()/len(x)*100,1)).reset_index()
        prod_fail.columns = ['제품명','불합격률(%)']
        prod_fail = prod_fail.sort_values('불합격률(%)', ascending=False)
        fig3 = px.bar(prod_fail, x='제품명', y='불합격률(%)',
                     title='제품별 불합격률 | Fail Rate by Product',
                     color='불합격률(%)', color_continuous_scale='Reds')
        fig3.update_layout(plot_bgcolor='white', height=280)
        st.plotly_chart(fig3, use_container_width=True)

    with c4:
        insp_perf = df.groupby('검사자').apply(
            lambda x: round((x['판정결과']=='합격').sum()/len(x)*100,1)).reset_index()
        insp_perf.columns = ['검사자','합격률(%)']
        fig4 = px.bar(insp_perf, x='검사자', y='합격률(%)',
                     title='검사자별 합격률 | Pass Rate by Inspector',
                     color='합격률(%)', color_continuous_scale='Blues')
        fig4.add_hline(y=90, line_dash="dash", line_color="orange")
        fig4.update_layout(plot_bgcolor='white', height=280)
        st.plotly_chart(fig4, use_container_width=True)

    fpath, fname = auto_save_excel(trend, '불량트렌드분석')
    if fpath:
        st.success(f'✅ 자동저장: 25_output_data/{fname}')

# ══════════════════════════════════════════════
# ⑧ AI 판정 & 인사이트
# ══════════════════════════════════════════════
elif menu == "🤖 AI 판정 & 인사이트":
    st.markdown('<div class="sec-title">🤖 AI 판정 & 인사이트 | AI Judgment & Insights</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 🧠 AI 종합 분석 | AI Comprehensive Analysis")
        total_ai   = len(df)
        fail_ai    = len(df[df['판정결과']=='불합격'])
        fail_rate  = round(fail_ai/total_ai*100, 1)
        top_defect = df[df['불량유형']!='없음']['불량유형'].value_counts().index[0]
        top_line   = df[df['판정결과']=='불합격']['라인'].value_counts().index[0]
        top_prod   = df[df['판정결과']=='불합격']['제품명'].value_counts().index[0]
        critical_r = round(len(df[df['불량등급']=='Critical'])/total_ai*100, 1)

        st.markdown(f"""
        <div class="info-box">
        🤖 <b>AI 분석 결과 요약 | AI Analysis Summary</b><br><br>
        📊 전체 불합격률: <b>{fail_rate}%</b><br>
        🔴 Critical 불량률: <b>{critical_r}%</b><br>
        🔍 최다 불량유형: <b>{top_defect}</b><br>
        🏭 불량 최다 라인: <b>{top_line}</b><br>
        📦 불량 최다 제품: <b>{top_prod}</b><br>
        📈 트렌드: {'⬆️ 불량 증가 추세 → 즉시 조치 필요' if fail_rate > 15 else '⬇️ 불량 안정 추세'}
        </div>""", unsafe_allow_html=True)

        # 불량등급별 파레토
        grade_cnt = df[df['불량등급']!='양호']['불량등급'].value_counts().reset_index()
        grade_cnt.columns = ['등급','건수']
        grade_cnt['누적비율(%)'] = (grade_cnt['건수'].cumsum() / grade_cnt['건수'].sum() * 100).round(1)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=grade_cnt['등급'], y=grade_cnt['건수'],
                            name='건수', marker_color='#1A3080'))
        fig.add_trace(go.Scatter(x=grade_cnt['등급'], y=grade_cnt['누적비율(%)'],
                                name='누적비율(%)', yaxis='y2',
                                line=dict(color='#FFD700', width=2), marker_size=8))
        fig.update_layout(
            title='불량등급 파레토 | Defect Grade Pareto',
            yaxis2=dict(overlaying='y', side='right', range=[0,110]),
            plot_bgcolor='white', height=300,
            legend=dict(orientation='h')
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("#### 💡 AI 개선 권고사항 | AI Recommendations")
        recommendations = [
            ("🔴 긴급", f"{top_line} 즉시 점검", f"불량 최다 발생 라인. 설비 점검 및 작업 표준 재확인 필요"),
            ("🟡 주의", f"{top_defect} 집중 관리", f"최다 불량유형. 작업 조건 및 원자재 품질 즉시 확인"),
            ("🟠 개선", f"{top_prod} 전수검사 실시", f"불량 최다 제품. 출하 전 전수검사로 유출 방지"),
            ("🟢 예방", "AI 검사 자동화 확대", f"현재 합격률 {100-fail_rate}%. AI 비전 검사로 99% 목표 달성 가능"),
            ("🔵 분석", "검사자 간 편차 분석", "검사자별 합격률 편차 존재. 검사 기준 교육 강화 필요"),
            ("🟣 혁신", "실시간 SPC 관리 도입", "통계적 공정관리로 불량 예측 및 사전 예방 체계 구축"),
        ]
        for priority, title, desc in recommendations:
            st.markdown(f"""
            <div style="background:white;border-radius:8px;padding:12px 16px;
                        margin-bottom:8px;box-shadow:0 2px 6px rgba(0,0,0,0.08);
                        border-left:4px solid #1A3080;">
            <b>{priority} {title}</b><br>
            <span style="font-size:12px;color:#555;">{desc}</span>
            </div>""", unsafe_allow_html=True)

        # 제품별 불량등급 히트맵
        heat_df = df.groupby(['제품명','불량등급']).size().reset_index(name='건수')
        heat_pivot = heat_df.pivot(index='제품명', columns='불량등급', values='건수').fillna(0)
        fig2 = px.imshow(heat_pivot, title='제품×등급 히트맵 | Product×Grade Heatmap',
                        color_continuous_scale='Blues', aspect='auto')
        fig2.update_layout(height=280)
        st.plotly_chart(fig2, use_container_width=True)

    rec_df = pd.DataFrame({
        '우선순위': [r[0] for r in recommendations],
        '제목':     [r[1] for r in recommendations],
        '내용':     [r[2] for r in recommendations],
    })
    fpath, fname = auto_save_excel(rec_df, 'AI인사이트')
    if fpath:
        st.success(f'✅ 자동저장: 25_output_data/{fname}')

# ──────────────────────────────────────────────
# 공통 푸터
# ──────────────────────────────────────────────
st.markdown("""
<div class="footer-bar">
  <b>경영시스템(주)</b> &nbsp;|&nbsp; AI 담당 최병대 &nbsp;|&nbsp; 010-3533-9030 &nbsp;|&nbsp;
  MSG-03 AI 기반 품질 검사 자동화 &nbsp;|&nbsp;
  AI program base 300ea &nbsp;|&nbsp; Ver 1.0 &nbsp;|&nbsp; 2026.03
</div>
""", unsafe_allow_html=True)
