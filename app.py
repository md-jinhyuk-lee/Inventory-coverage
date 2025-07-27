import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import warnings

warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="재고 커버리지 분석 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def convert_df_to_excel(df, sheet_name='Sheet1'):
    """DataFrame을 Excel로 변환"""
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return output.getvalue()
    except ImportError:
        st.error("❌ Excel 기능을 사용하려면 openpyxl 라이브러리가 필요합니다.")
        st.info("터미널에서 다음 명령어를 실행하세요: pip install openpyxl")
        return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    except Exception as e:
        st.error(f"Excel 변환 중 오류: {str(e)}")
        return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

def create_html_report(data):
    """HTML 리포트 생성"""
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            .metrics {{ display: flex; justify-content: space-around; margin: 20px 0; }}
            .metric {{ background-color: #e8f4fd; padding: 15px; text-align: center; border-radius: 8px; margin: 5px; }}
            .metric h3 {{ margin: 0; color: #1f77b4; }}
            .metric p {{ margin: 5px 0 0 0; font-size: 18px; font-weight: bold; }}
            table {{ width: 100%; border-collapse: collapse; margin: 10px 0; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
            .section {{ margin: 20px 0; }}
        </style>
    </head>
    <body>
        <h1>📊 재고 커버리지 분석 리포트</h1>
        <p>보고서 생성일: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
        
        <h2>📊 전체 요약</h2>
        <div class="metrics">
            <div class="metric">
                <h3>총 매장 수</h3>
                <p>{len(data['매장명'].unique())}개</p>
            </div>
            <div class="metric">
                <h3>총 상품 수</h3>
                <p>{len(data['상품코드'].unique())}개</p>
            </div>
            <div class="metric">
                <h3>위험상품 수</h3>
                <p>{len(data[data['status'] == 'critical'])}개</p>
            </div>
            <div class="metric">
                <h3>평균 커버리지</h3>
                <p>{data['coverage_weeks'].mean():.1f}주</p>
            </div>
        </div>
        
        <h2>📄 종합 리포트 주요 지표</h2>
        <div class="metrics">
            <div class="metric">
                <h3>총 재고 금액</h3>
                <p>{data['재고_금액'].sum():,.0f}원</p>
            </div>
            <div class="metric">
                <h3>주간 예상 매출</h3>
                <p>{(data['소비자가'] * data['avg_weekly_sales']).sum():,.0f}원</p>
            </div>
            <div class="metric">
                <h3>위험상품 비율</h3>
                <p>{len(data[data['status'] == 'critical']) / len(data) * 100:.1f}%</p>
            </div>
            <div class="metric">
                <h3>전체 평균 커버리지</h3>
                <p>{data['coverage_weeks'].mean():.1f}주</p>
            </div>
        </div>
    </body>
    </html>
    """
    return html_content

def send_email_report(data, recipient_email, sender_email, sender_password):
    """이메일 리포트 발송 - 첨부파일 제외"""
    try:
        msg = MIMEMultipart('mixed')
        msg['Subject'] = f"재고 커버리지 분석 리포트 - {datetime.now().strftime('%Y-%m-%d')}"
        msg['From'] = sender_email
        msg['To'] = recipient_email
        
        # BIZ별 전체요약 데이터 준비
        biz_summary = []
        biz_order = ['AP', 'FW', 'EQ']
        for biz in biz_order:
            if biz in data['BIZ'].values:
                biz_data = data[data['BIZ'] == biz]
                total_sales = biz_data['1주차_판매량'].sum() + biz_data['2주차_판매량'].sum() + biz_data['3주차_판매량'].sum()
                biz_summary.append({
                    'BIZ': biz,
                    '총_매장_수': len(biz_data['매장명'].unique()),
                    '총_상품_수': len(biz_data['상품코드'].unique()),
                    '위험상품_수': len(biz_data[biz_data['status'] == 'critical']),
                    '평균_커버리지': f"{biz_data['coverage_weeks'].mean():.1f}주",
                    '판매수량': int(total_sales),
                    '재고수량': int(biz_data['현재_재고량'].sum()),
                    '재고금액': f"{int(biz_data['재고_금액'].sum()):,}원"
                })
        
        # BIZ별 TOTAL 추가
        total_sales_all = data['1주차_판매량'].sum() + data['2주차_판매량'].sum() + data['3주차_판매량'].sum()
        biz_summary.append({
            'BIZ': 'TOTAL',
            '총_매장_수': len(data['매장명'].unique()),
            '총_상품_수': len(data['상품코드'].unique()),
            '위험상품_수': len(data[data['status'] == 'critical']),
            '평균_커버리지': f"{data['coverage_weeks'].mean():.1f}주",
            '판매수량': int(total_sales_all),
            '재고수량': int(data['현재_재고량'].sum()),
            '재고금액': f"{int(data['재고_금액'].sum()):,}원"
        })
        
        # 시즌별 전체요약 데이터 준비
        season_summary = []
        for season in sorted(data['시즌'].unique()):
            season_data = data[data['시즌'] == season]
            season_sales = season_data['1주차_판매량'].sum() + season_data['2주차_판매량'].sum() + season_data['3주차_판매량'].sum()
            season_summary.append({
                '시즌': season,
                '총_매장_수': len(season_data['매장명'].unique()),
                '총_상품_수': len(season_data['상품코드'].unique()),
                '위험상품_수': len(season_data[season_data['status'] == 'critical']),
                '평균_커버리지': f"{season_data['coverage_weeks'].mean():.1f}주",
                '판매수량': int(season_sales),
                '재고수량': int(season_data['현재_재고량'].sum()),
                '재고금액': f"{int(season_data['재고_금액'].sum()):,}원"
            })
        
        # 시즌별 TOTAL 추가
        season_summary.append({
            '시즌': 'TOTAL',
            '총_매장_수': len(data['매장명'].unique()),
            '총_상품_수': len(data['상품코드'].unique()),
            '위험상품_수': len(data[data['status'] == 'critical']),
            '평균_커버리지': f"{data['coverage_weeks'].mean():.1f}주",
            '판매수량': int(total_sales_all),
            '재고수량': int(data['현재_재고량'].sum()),
            '재고금액': f"{int(data['재고_금액'].sum()):,}원"
        })
        
        # BIZ별 종합리포트 데이터 준비
        biz_report = []
        critical_by_biz = data[data['status'] == 'critical'].groupby('BIZ').size()
        for biz in biz_order:
            if biz in data['BIZ'].values:
                biz_data = data[data['BIZ'] == biz]
                critical_count = critical_by_biz.get(biz, 0)
                biz_report.append({
                    'BIZ': biz,
                    '총_재고_금액': f"{biz_data['재고_금액'].sum():,.0f}원",
                    '주간_예상_매출': f"{(biz_data['소비자가'] * biz_data['avg_weekly_sales']).sum():,.0f}원",
                    '위험상품_비율': f"{(critical_count / len(biz_data) * 100) if len(biz_data) > 0 else 0:.1f}%",
                    '전체_평균_커버리지': f"{biz_data['coverage_weeks'].mean():.1f}주"
                })
        
        # 종합리포트 TOTAL 추가
        total_critical = len(data[data['status'] == 'critical'])
        biz_report.append({
            'BIZ': 'TOTAL',
            '총_재고_금액': f"{data['재고_금액'].sum():,.0f}원",
            '주간_예상_매출': f"{(data['소비자가'] * data['avg_weekly_sales']).sum():,.0f}원",
            '위험상품_비율': f"{(total_critical / len(data) * 100):.1f}%",
            '전체_평균_커버리지': f"{data['coverage_weeks'].mean():.1f}주"
        })
        
        # BIZ별 전체요약 테이블 HTML
        biz_summary_table = """
        <table style="width: 100%; border-collapse: collapse; margin: 10px 0;">
            <thead>
                <tr style="background-color: #f2f2f2;">
                    <th style="border: 1px solid #ddd; padding: 8px;">BIZ</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">총 매장 수</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">총 상품 수</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">위험상품 수</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">평균 커버리지</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">판매수량</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">재고수량</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">재고금액</th>
                </tr>
            </thead>
            <tbody>
        """
        for item in biz_summary:
            style = "background-color: #000000; color: white; font-weight: bold;" if item['BIZ'] == 'TOTAL' else ""
            biz_summary_table += f"""
                <tr style="{style}">
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['BIZ']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['총_매장_수']}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['총_상품_수']}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['위험상품_수']}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['평균_커버리지']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['판매수량']:,}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['재고수량']:,}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['재고금액']}</td>
                </tr>
            """
        biz_summary_table += "</tbody></table>"
        
        # 시즌별 전체요약 테이블 HTML
        season_summary_table = """
        <table style="width: 100%; border-collapse: collapse; margin: 10px 0;">
            <thead>
                <tr style="background-color: #f2f2f2;">
                    <th style="border: 1px solid #ddd; padding: 8px;">시즌</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">총 매장 수</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">총 상품 수</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">위험상품 수</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">평균 커버리지</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">판매수량</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">재고수량</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">재고금액</th>
                </tr>
            </thead>
            <tbody>
        """
        for item in season_summary:
            style = "background-color: #000000; color: white; font-weight: bold;" if item['시즌'] == 'TOTAL' else ""
            season_summary_table += f"""
                <tr style="{style}">
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['시즌']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['총_매장_수']}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['총_상품_수']}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['위험상품_수']}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['평균_커버리지']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['판매수량']:,}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['재고수량']:,}개</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['재고금액']}</td>
                </tr>
            """
        season_summary_table += "</tbody></table>"
        
        # BIZ별 종합리포트 테이블 HTML
        biz_report_table = """
        <table style="width: 100%; border-collapse: collapse; margin: 10px 0;">
            <thead>
                <tr style="background-color: #f2f2f2;">
                    <th style="border: 1px solid #ddd; padding: 8px;">BIZ</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">총 재고 금액</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">주간 예상 매출</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">위험상품 비율</th>
                    <th style="border: 1px solid #ddd; padding: 8px;">전체 평균 커버리지</th>
                </tr>
            </thead>
            <tbody>
        """
        for item in biz_report:
            style = "background-color: #000000; color: white; font-weight: bold;" if item['BIZ'] == 'TOTAL' else ""
            biz_report_table += f"""
                <tr style="{style}">
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['BIZ']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['총_재고_금액']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['주간_예상_매출']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['위험상품_비율']}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{item['전체_평균_커버리지']}</td>
                </tr>
            """
        biz_report_table += "</tbody></table>"
        
        # 이메일 본문 HTML
        html_content = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                .metrics {{ display: flex; justify-content: space-around; margin: 20px 0; }}
                .metric {{ background-color: #e8f4fd; padding: 15px; text-align: center; border-radius: 8px; margin: 5px; }}
                .metric h3 {{ margin: 0; color: #1f77b4; }}
                .metric p {{ margin: 5px 0 0 0; font-size: 18px; font-weight: bold; }}
                table {{ width: 100%; border-collapse: collapse; margin: 10px 0; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .section {{ margin: 20px 0; }}
            </style>
        </head>
        <body>
            <h1>📊 재고 커버리지 분석 리포트</h1>
            <p>보고서 생성일: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
            
            <h2>📊 전체 요약 - BIZ별 구분</h2>
            {biz_summary_table}
            
            <h2>📊 전체 요약 - 시즌별 구분</h2>
            {season_summary_table}
            
            <h2>📄 종합 리포트 주요 지표 - BIZ별 구분</h2>
            {biz_report_table}
            
            <p><strong>📋 재고 상태 분류 기준:</strong></p>
            <ul>
                <li>🚨 위험: 2주 미만 (즉시 보충 필요)</li>
                <li>⚠️ 주의: 2주 이상 ~ 4주 미만 (보충 검토 필요)</li>
                <li>✅ 양호: 4주 이상 (안정적인 재고 수준)</li>
            </ul>
            
            <p>상세한 분석 결과는 대시보드에서 확인해주세요.</p>
        </body>
        </html>
        """
        
        html_part = MIMEText(html_content, 'html', 'utf-8')
        msg.attach(html_part)
        
        # 이메일 발송
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        
        return True, "이메일이 성공적으로 발송되었습니다! 📧"
    
    except Exception as e:
        return False, f"이메일 발송 실패: {str(e)}"

def load_and_process_data(uploaded_file):
    """데이터 로드 및 처리"""
    try:
        data = pd.read_excel(uploaded_file)
        data.columns = data.columns.str.strip()
        
        required_columns = ['매장명', '상품명', '상품코드', 'BIZ', '시즌', '소비자가', 
                           '1주차_판매량', '2주차_판매량', '3주차_판매량', '현재_재고량', '재고_금액']
        
        missing_columns = [col for col in required_columns if col not in data.columns]
        if missing_columns:
            return None, f"필수 컬럼이 누락되었습니다: {missing_columns}"
        
        data = data.dropna(subset=['매장명', '상품명', '상품코드'])
        
        # 숫자 컬럼 변환
        numeric_columns = ['소비자가', '1주차_판매량', '2주차_판매량', '3주차_판매량', '현재_재고량', '재고_금액']
        for col in numeric_columns:
            if data[col].dtype == 'object':
                data[col] = data[col].astype(str).str.replace(',', '').str.replace('원', '').str.replace(' ', '')
            data[col] = pd.to_numeric(data[col], errors='coerce')
        
        data = data.dropna(subset=numeric_columns)
        
        # 분석 컬럼 추가
        data['avg_weekly_sales'] = (data['1주차_판매량'] + data['2주차_판매량'] + data['3주차_판매량']) / 3
        data['coverage_weeks'] = np.where(
            data['avg_weekly_sales'] == 0,
            999,
            data['현재_재고량'] / data['avg_weekly_sales']
        )
        
        def classify_status(weeks):
            if weeks < 2:
                return 'critical'
            elif weeks < 4:
                return 'warning'
            else:
                return 'good'
        
        data['status'] = data['coverage_weeks'].apply(classify_status)
        
        return data, None
        
    except Exception as e:
        return None, f"데이터 처리 중 오류: {str(e)}"

# 색상 팔레트 정의
SEASON_COLORS = ['#FFE5E5', '#E5F3FF', '#E5FFE5', '#FFF5E5', '#F0E5FF', '#FFE5F5']
BIZ_COLORS = ['#E8F4FD', '#FFF2E8', '#E8F5E8', '#F5E8F5', '#E8E8F5']
STORE_COLORS = ['#FFE5E5', '#E5F3FF', '#E5FFE5', '#FFF5E5', '#F0E5FF', '#FFE5F5', '#E5FFFF', '#FFE5DD', '#E5E5FF', '#F5FFE5']

# 메인 앱 시작
st.title("📊 재고 커버리지 분석 대시보드")

st.markdown("""
**📋 재고 상태 분류 기준:**
- 🚨 **위험**: 2주 미만 (즉시 보충 필요)  
- ⚠️ **주의**: 2주 이상 ~ 4주 미만 (보충 검토 필요)  
- ✅ **양호**: 4주 이상 (안정적인 재고 수준)

**계산 방식:** 현재 재고량 ÷ 3주 평균 판매량 = 재고 커버리지(주)
""")

st.markdown("---")

# 사이드바
with st.sidebar:
    st.header("📁 데이터 업로드")
    uploaded_file = st.file_uploader(
        "Excel 파일을 업로드하세요",
        type=['xlsx', 'xls']
    )

if uploaded_file is not None:
    # 데이터 로드
    data, error_msg = load_and_process_data(uploaded_file)
    
    if error_msg:
        st.error(f"❌ {error_msg}")
        if data is None:
            st.write("📋 현재 파일의 컬럼들:")
            try:
                temp_data = pd.read_excel(uploaded_file)
                st.write(list(temp_data.columns))
            except Exception:
                st.write("파일을 읽을 수 없습니다.")
    else:
        st.success("✅ 데이터 로드 완료!")
        
        # 메뉴
        with st.sidebar:
            st.header("📋 메뉴")
            menu = st.radio("선택하세요", [
                "📊 전체 요약",
                "🏢 BIZ별 분석",
                "🌸 시즌별 분석", 
                "🏪 매장별 상세 분석",
                "🔍 상세 분석",
                "📄 종합 리포트",
                "📧 이메일 발송"
            ])
        
        # 메뉴별 실행
        if menu == "📊 전체 요약":
            st.header("📊 전체 요약")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("총 매장 수", len(data['매장명'].unique()))
            
            with col2:
                st.metric("총 상품 수", len(data['상품코드'].unique()))
            
            with col3:
                critical_count = len(data[data['status'] == 'critical'])
                st.metric("위험상품 수", critical_count)
            
            with col4:
                avg_coverage = round(data['coverage_weeks'].mean(), 1)
                st.metric("평균 커버리지", f"{avg_coverage}주")
            
            # 시즌별 재고 상태 분포
            st.subheader("📊 시즌별 재고 상태 분포")
            season_status = data.groupby(['시즌', 'status']).size().reset_index(name='count')
            season_status['status_korean'] = season_status['status'].map({'critical': '위험', 'warning': '주의', 'good': '양호'})
            
            fig1 = px.bar(
                season_status,
                x='시즌',
                y='count',
                color='status_korean',
                title="시즌별 재고 상태 분포",
                labels={'count': '상품 수', 'status_korean': '재고 상태'},
                color_discrete_map={'위험': '#e74c3c', '주의': '#f39c12', '양호': '#27ae60'},
                text='count'
            )
            fig1.update_traces(textposition='outside', textfont_size=16)
            fig1.update_layout(yaxis=dict(range=[0, 9000]))
            st.plotly_chart(fig1, use_container_width=True)
            
            # 시즌별 평균 커버리지
            st.subheader("📈 시즌별 평균 커버리지")
            season_coverage = data.groupby('시즌')['coverage_weeks'].mean().reset_index()
            season_coverage.columns = ['시즌', '평균_커버리지']
            
            fig2 = px.bar(
                season_coverage,
                x='시즌',
                y='평균_커버리지',
                title="시즌별 평균 커버리지",
                text='평균_커버리지',
                color='시즌',
                color_discrete_sequence=SEASON_COLORS
            )
            fig2.update_traces(texttemplate='%{text:.1f}주', textposition='outside', textfont_size=16)
            fig2.update_layout(yaxis=dict(range=[0, 70]))
            st.plotly_chart(fig2, use_container_width=True)
            
            # BIZ별 재고 상태 분포
            st.subheader("🏢 BIZ별 재고 상태 분포")
            biz_status = data.groupby(['BIZ', 'status']).size().reset_index(name='count')
            biz_status['status_korean'] = biz_status['status'].map({'critical': '위험', 'warning': '주의', 'good': '양호'})
            
            # BIZ 순서 정렬 (AP, FW, EQ 순)
            biz_order = ['AP', 'FW', 'EQ']
            if len(set(biz_order) & set(biz_status['BIZ'].unique())) > 0:
                biz_status['BIZ'] = pd.Categorical(biz_status['BIZ'], categories=biz_order, ordered=True)
                biz_status = biz_status.sort_values('BIZ')
            
            fig3 = px.bar(
                biz_status,
                x='BIZ',
                y='count',
                color='status_korean',
                title="BIZ별 재고 상태 분포",
                labels={'count': '상품 수', 'status_korean': '재고 상태'},
                color_discrete_map={'위험': '#e74c3c', '주의': '#f39c12', '양호': '#27ae60'},
                text='count'
            )
            fig3.update_traces(textposition='outside', textfont_size=16)
            fig3.update_layout(yaxis=dict(range=[0, 9000]))
            st.plotly_chart(fig3, use_container_width=True)
            
            # BIZ별 평균 커버리지
            st.subheader("📈 BIZ별 평균 커버리지")
            biz_coverage = data.groupby('BIZ')['coverage_weeks'].mean().reset_index()
            biz_coverage.columns = ['BIZ', '평균_커버리지']
            
            # BIZ 순서 정렬 (AP, FW, EQ 순)
            if len(set(biz_order) & set(biz_coverage['BIZ'].unique())) > 0:
                biz_coverage['BIZ'] = pd.Categorical(biz_coverage['BIZ'], categories=biz_order, ordered=True)
                biz_coverage = biz_coverage.sort_values('BIZ')
            
            fig4 = px.bar(
                biz_coverage,
                x='BIZ',
                y='평균_커버리지',
                title="BIZ별 평균 커버리지",
                text='평균_커버리지',
                color='BIZ',
                color_discrete_sequence=BIZ_COLORS
            )
            fig4.update_traces(texttemplate='%{text:.1f}주', textposition='outside', textfont_size=16)
            fig4.update_layout(yaxis=dict(range=[0, 30]))
            st.plotly_chart(fig4, use_container_width=True)
            
            # 온라인 제외 필터링
            offline_data = data[data['매장명'] != '온라인']
            
            # 매장별 분석 데이터 생성
            store_analysis = []
            for store in offline_data['매장명'].unique():
                store_data = offline_data[offline_data['매장명'] == store]
                store_analysis.append({
                    '매장명': store,
                    '평균_커버리지': round(store_data['coverage_weeks'].mean(), 1),
                    '총_상품코드수': len(store_data['상품코드'].unique()),
                    '위험상품수': len(store_data[store_data['status'] == 'critical']),
                    '주의상품수': len(store_data[store_data['status'] == 'warning']),
                    '양호상품수': len(store_data[store_data['status'] == 'good']),
                    '재고_금액': store_data['재고_금액'].sum()
                })
            
            store_df = pd.DataFrame(store_analysis)
            
            # 매출이 높은 순으로 정렬하여 2-6위 가져오기
            store_sales_ranking = store_df.sort_values('재고_금액', ascending=False)
            
            # 양호 상위 5개 매장 (상위 2번째부터 6번째)
            st.subheader("✅ 양호 상위 5개 매장")
            if len(store_sales_ranking) >= 6:
                top_stores_2to6 = store_sales_ranking.iloc[1:6]  # 2번째부터 6번째
                good_stores = top_stores_2to6.sort_values('양호상품수', ascending=False)
                good_stores_display = good_stores[['매장명', '평균_커버리지', '총_상품코드수', '양호상품수']].copy()
                st.dataframe(good_stores_display, use_container_width=True, hide_index=True)
            else:
                good_stores = store_df.nlargest(5, '양호상품수')
                good_stores_display = good_stores[['매장명', '평균_커버리지', '총_상품코드수', '양호상품수']].copy()
                st.dataframe(good_stores_display, use_container_width=True, hide_index=True)
            
            # 주의 상위 5개 매장
            st.subheader("⚠️ 주의 상위 5개 매장")
            warning_stores = store_df.nlargest(5, '주의상품수')
            warning_stores_display = warning_stores[['매장명', '평균_커버리지', '총_상품코드수', '주의상품수']].copy()
            st.dataframe(warning_stores_display, use_container_width=True, hide_index=True)
            
            # 위험 상위 5개 매장
            st.subheader("🚨 위험 상위 5개 매장")
            critical_stores = store_df.nlargest(5, '위험상품수')
            critical_stores_display = critical_stores[['매장명', '평균_커버리지', '총_상품코드수', '위험상품수']].copy()
            st.dataframe(critical_stores_display, use_container_width=True, hide_index=True)
        
        elif menu == "🏢 BIZ별 분석":
            st.header("🏢 BIZ별 분석")
            
            # BIZ 순서: AP, FW, EQ
            biz_order = ['AP', 'FW', 'EQ']
            available_biz = [biz for biz in biz_order if biz in data['BIZ'].unique()]
            other_biz = [biz for biz in sorted(data['BIZ'].unique()) if biz not in biz_order]
            all_biz = available_biz + other_biz
            
            # BIZ별 분석 테이블
            biz_analysis = []
            for biz in all_biz:
                biz_data = data[data['BIZ'] == biz]
                biz_analysis.append({
                    'BIZ': biz,
                    '총_상품수': len(biz_data),
                    '평균_커버리지': round(biz_data['coverage_weeks'].mean(), 1),
                    '위험상품수': len(biz_data[biz_data['status'] == 'critical']),
                    '주의상품수': len(biz_data[biz_data['status'] == 'warning']),
                    '양호상품수': len(biz_data[biz_data['status'] == 'good']),
                    '재고_수량': int(biz_data['현재_재고량'].sum()),
                    '재고_금액': f"{int(biz_data['재고_금액'].sum()):,}원"
                })
            
            # TOTAL 행 추가
            biz_analysis.append({
                'BIZ': 'TOTAL',
                '총_상품수': len(data),
                '평균_커버리지': round(data['coverage_weeks'].mean(), 1),
                '위험상품수': len(data[data['status'] == 'critical']),
                '주의상품수': len(data[data['status'] == 'warning']),
                '양호상품수': len(data[data['status'] == 'good']),
                '재고_수량': int(data['현재_재고량'].sum()),
                '재고_금액': f"{int(data['재고_금액'].sum()):,}원"
            })
            
            biz_df = pd.DataFrame(biz_analysis)
            
            # TOTAL 행 하이라이트
            def highlight_total(row):
                if row['BIZ'] == 'TOTAL':
                    return ['background-color: #000000; color: white; font-weight: bold;'] * len(row)
                return [''] * len(row)
            
            styled_biz = biz_df.style.apply(highlight_total, axis=1)
            st.dataframe(styled_biz, use_container_width=True, hide_index=True)
            
            # BIZ별 차트
            col1, col2 = st.columns(2)
            
            with col1:
                biz_status = data.groupby(['BIZ', 'status']).size().reset_index(name='count')
                biz_status['status_korean'] = biz_status['status'].map({'critical': '위험', 'warning': '주의', 'good': '양호'})
                
                # BIZ 순서 정렬 (AP, FW, EQ 순)
                if len(set(biz_order) & set(biz_status['BIZ'].unique())) > 0:
                    biz_status['BIZ'] = pd.Categorical(biz_status['BIZ'], categories=biz_order, ordered=True)
                    biz_status = biz_status.sort_values('BIZ')
                
                fig1 = px.bar(
                    biz_status,
                    x='BIZ',
                    y='count',
                    color='status_korean',
                    title="BIZ별 재고 상태 분포",
                    labels={'count': '상품 수', 'status_korean': '재고 상태'},
                    color_discrete_map={'위험': '#e74c3c', '주의': '#f39c12', '양호': '#27ae60'},
                    text='count'
                )
                fig1.update_traces(textposition='outside', textfont_size=16)
                fig1.update_layout(yaxis=dict(range=[0, 9000]))
                st.plotly_chart(fig1, use_container_width=True)
            
            with col2:
                biz_coverage = data.groupby('BIZ')['coverage_weeks'].mean().reset_index()
                biz_coverage.columns = ['BIZ', '평균_커버리지']
                
                # BIZ 순서 정렬 (AP, FW, EQ 순)
                if len(set(biz_order) & set(biz_coverage['BIZ'].unique())) > 0:
                    biz_coverage['BIZ'] = pd.Categorical(biz_coverage['BIZ'], categories=biz_order, ordered=True)
                    biz_coverage = biz_coverage.sort_values('BIZ')
                
                fig2 = px.bar(
                    biz_coverage,
                    x='BIZ',
                    y='평균_커버리지',
                    title="BIZ별 평균 커버리지",
                    text='평균_커버리지',
                    color='BIZ',
                    color_discrete_sequence=BIZ_COLORS
                )
                fig2.update_traces(texttemplate='%{text:.1f}주', textposition='outside', textfont_size=16)
                fig2.update_layout(yaxis=dict(range=[0, 30]))
                st.plotly_chart(fig2, use_container_width=True)
        
        elif menu == "🌸 시즌별 분석":
            st.header("🌸 시즌별 분석")
            
            # 시즌별 분석 테이블
            season_analysis = []
            for season in sorted(data['시즌'].unique()):
                season_data = data[data['시즌'] == season]
                season_analysis.append({
                    '시즌': season,
                    '총_상품수': len(season_data),
                    '평균_커버리지': round(season_data['coverage_weeks'].mean(), 1),
                    '위험상품수': len(season_data[season_data['status'] == 'critical']),
                    '주의상품수': len(season_data[season_data['status'] == 'warning']),
                    '양호상품수': len(season_data[season_data['status'] == 'good']),
                    '재고_수량': int(season_data['현재_재고량'].sum()),
                    '재고_금액': f"{int(season_data['재고_금액'].sum()):,}원"
                })
            
            # TOTAL 행 추가
            season_analysis.append({
                '시즌': 'TOTAL',
                '총_상품수': len(data),
                '평균_커버리지': round(data['coverage_weeks'].mean(), 1),
                '위험상품수': len(data[data['status'] == 'critical']),
                '주의상품수': len(data[data['status'] == 'warning']),
                '양호상품수': len(data[data['status'] == 'good']),
                '재고_수량': int(data['현재_재고량'].sum()),
                '재고_금액': f"{int(data['재고_금액'].sum()):,}원"
            })
            
            season_df = pd.DataFrame(season_analysis)
            
            # TOTAL 행 하이라이트
            def highlight_total_season(row):
                if row['시즌'] == 'TOTAL':
                    return ['background-color: #000000; color: white; font-weight: bold;'] * len(row)
                return [''] * len(row)
            
            styled_season = season_df.style.apply(highlight_total_season, axis=1)
            st.dataframe(styled_season, use_container_width=True, hide_index=True)
            
            # 시즌별 차트
            col1, col2 = st.columns(2)
            
            with col1:
                season_status = data.groupby(['시즌', 'status']).size().reset_index(name='count')
                season_status['status_korean'] = season_status['status'].map({'critical': '위험', 'warning': '주의', 'good': '양호'})
                
                fig1 = px.bar(
                    season_status,
                    x='시즌',
                    y='count',
                    color='status_korean',
                    title="시즌별 재고 상태 분포",
                    labels={'count': '상품 수', 'status_korean': '재고 상태'},
                    color_discrete_map={'위험': '#e74c3c', '주의': '#f39c12', '양호': '#27ae60'},
                    text='count'
                )
                fig1.update_traces(textposition='outside', textfont_size=16)
                fig1.update_layout(yaxis=dict(range=[0, 9000]))
                st.plotly_chart(fig1, use_container_width=True)
            
            with col2:
                season_coverage = data.groupby('시즌')['coverage_weeks'].mean().reset_index()
                season_coverage.columns = ['시즌', '평균_커버리지']
                
                fig2 = px.bar(
                    season_coverage,
                    x='시즌',
                    y='평균_커버리지',
                    title="시즌별 평균 커버리지",
                    text='평균_커버리지',
                    color='시즌',
                    color_discrete_sequence=SEASON_COLORS
                )
                fig2.update_traces(texttemplate='%{text:.1f}주', textposition='outside', textfont_size=16)
                fig2.update_layout(yaxis=dict(range=[0, 70]))
                st.plotly_chart(fig2, use_container_width=True)
        
        elif menu == "🏪 매장별 상세 분석":
            st.header("🏪 매장별 상세 분석")
            
            # 매장 선택 - 온라인을 제일 마지막에 배치
            offline_stores = sorted([store for store in data['매장명'].unique() if store != '온라인'])
            online_stores = [store for store in data['매장명'].unique() if store == '온라인']
            store_list = ['전체'] + offline_stores + online_stores
            
            selected_stores = st.multiselect(
                "분석할 매장을 선택하세요:",
                store_list,
                default=['전체']
            )
            
            if '전체' in selected_stores:
                # 온라인 제외한 전체 매장 데이터
                offline_stores_data = data[data['매장명'] != '온라인']
                store_sales = offline_stores_data.groupby('매장명')['재고_금액'].sum().sort_values(ascending=False)
                top_stores = store_sales.head(11).index.tolist()  # 상위 11개 가져와서 2-11위 사용
                
                # 매장별 재고 상태 분포
                st.subheader("📊 매장별 재고 상태 분포")
                top_store_data = offline_stores_data[offline_stores_data['매장명'].isin(top_stores[1:11])]  # 2번째부터 11번째
                top_store_status = top_store_data.groupby(['매장명', 'status']).size().reset_index(name='count')
                top_store_status['status_korean'] = top_store_status['status'].map({'critical': '위험', 'warning': '주의', 'good': '양호'})
                
                fig1 = px.bar(
                    top_store_status,
                    x='매장명',
                    y='count',
                    color='status_korean',
                    title="매장별 재고 상태 분포",
                    labels={'count': '상품 수', 'status_korean': '재고 상태'},
                    color_discrete_map={'위험': '#e74c3c', '주의': '#f39c12', '양호': '#27ae60'},
                    text='count'
                )
                fig1.update_traces(textposition='outside', textfont_size=16)
                fig1.update_layout(xaxis_tickangle=-45, yaxis=dict(range=[0, 1000]))
                st.plotly_chart(fig1, use_container_width=True)
                
                # 매장별 평균 커버리지
                st.subheader("📈 매장별 평균 커버리지")
                top_store_coverage = top_store_data.groupby('매장명')['coverage_weeks'].mean().reset_index()
                top_store_coverage.columns = ['매장명', '평균_커버리지']
                
                fig2 = px.bar(
                    top_store_coverage,
                    x='매장명',
                    y='평균_커버리지',
                    title="매장별 평균 커버리지",
                    text='평균_커버리지',
                    color='매장명',
                    color_discrete_sequence=STORE_COLORS
                )
                fig2.update_traces(texttemplate='%{text:.1f}주', textposition='outside', textfont_size=16)
                fig2.update_layout(xaxis_tickangle=-45, yaxis=dict(range=[0, 12]))
                st.plotly_chart(fig2, use_container_width=True)
                
                # 매장별 상세 분석 추가
                st.subheader("📋 매장별 상세 분석")
                store_analysis = []
                for store in offline_stores_data['매장명'].unique():
                    store_data = offline_stores_data[offline_stores_data['매장명'] == store]
                    store_analysis.append({
                        '매장명': store,
                        '평균_커버리지': round(store_data['coverage_weeks'].mean(), 1),
                        '총_상품코드수': len(store_data['상품코드'].unique()),
                        '위험상품수': len(store_data[store_data['status'] == 'critical']),
                        '주의상품수': len(store_data[store_data['status'] == 'warning']),
                        '양호상품수': len(store_data[store_data['status'] == 'good']),
                        '재고_수량': int(store_data['현재_재고량'].sum()),
                        '재고_금액': f"{int(store_data['재고_금액'].sum()):,}원"
                    })
                
                store_df = pd.DataFrame(store_analysis)
                st.dataframe(store_df, use_container_width=True, hide_index=True)
                
            elif selected_stores and '전체' not in selected_stores:
                filtered_data = data[data['매장명'].isin(selected_stores)]
                
                # 매장별 차트
                col1, col2 = st.columns(2)
                
                with col1:
                    store_status = filtered_data.groupby(['매장명', 'status']).size().reset_index(name='count')
                    store_status['status_korean'] = store_status['status'].map({'critical': '위험', 'warning': '주의', 'good': '양호'})
                    
                    fig1 = px.bar(
                        store_status,
                        x='매장명',
                        y='count',
                        color='status_korean',
                        title="매장별 재고 상태 분포",
                        labels={'count': '상품 수', 'status_korean': '재고 상태'},
                        color_discrete_map={'위험': '#e74c3c', '주의': '#f39c12', '양호': '#27ae60'},
                        text='count'
                    )
                    fig1.update_traces(textposition='outside', textfont_size=16)
                    fig1.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig1, use_container_width=True)
                
                with col2:
                    store_coverage = filtered_data.groupby('매장명')['coverage_weeks'].mean().reset_index()
                    store_coverage.columns = ['매장명', '평균_커버리지']
                    
                    fig2 = px.bar(
                        store_coverage,
                        x='매장명',
                        y='평균_커버리지',
                        title="매장별 평균 커버리지",
                        text='평균_커버리지',
                        color='매장명',
                        color_discrete_sequence=STORE_COLORS
                    )
                    fig2.update_traces(texttemplate='%{text:.1f}주', textposition='outside', textfont_size=16)
                    fig2.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig2, use_container_width=True)
            else:
                st.warning("⚠️ 매장을 선택해주세요.")
        
        elif menu == "🔍 상세 분석":
            st.header("🔍 상세 분석")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📊 전체 재고 상태 분포")
                status_counts = data['status'].value_counts()
                status_korean = {'critical': '위험', 'warning': '주의', 'good': '양호'}
                status_counts.index = [status_korean[idx] for idx in status_counts.index]
                
                fig_pie = px.pie(
                    values=status_counts.values,
                    names=status_counts.index,
                    title="전체 재고 상태 분포",
                    color_discrete_map={'위험': '#e74c3c', '주의': '#f39c12', '양호': '#27ae60'}
                )
                fig_pie.update_traces(textposition='inside', textinfo='percent+label+value', textfont_size=16)
                st.plotly_chart(fig_pie, use_container_width=True)
            
            with col2:
                st.subheader("📈 커버리지 분포 트리맵")
                coverage_data = data[data['coverage_weeks'] < 999]
                
                # 커버리지 구간별 집계
                coverage_bins = pd.cut(coverage_data['coverage_weeks'], bins=[0, 2, 4, 8, 12, 16, 20, 999], 
                                     labels=['0-2주', '2-4주', '4-8주', '8-12주', '12-16주', '16-20주', '20주+'])
                coverage_grouped = coverage_data.groupby(coverage_bins).size().reset_index()
                coverage_grouped.columns = ['커버리지_구간', 'SKU_수']
                coverage_grouped = coverage_grouped[coverage_grouped['SKU_수'] > 0]
                
                fig_treemap = px.treemap(
                    coverage_grouped,
                    path=['커버리지_구간'],
                    values='SKU_수',
                    title="커버리지 분포 트리맵",
                    color='SKU_수',
                    color_continuous_scale='Blues'
                )
                fig_treemap.update_traces(textinfo='label+value', textfont_size=16)
                st.plotly_chart(fig_treemap, use_container_width=True)
            
            # 상품별 상세 분석 (매장명 제외)
            st.subheader("🔍 상품별 상세 분석")
            
            # 필터링 옵션
            col1, col2, col3 = st.columns(3)
            
            with col1:
                selected_season = st.selectbox(
                    "시즌 선택:",
                    ['전체'] + sorted(data['시즌'].unique())
                )
            
            with col2:
                selected_biz = st.selectbox(
                    "BIZ 선택:",
                    ['전체'] + sorted(data['BIZ'].unique())
                )
            
            with col3:
                selected_status = st.selectbox(
                    "재고 상태 선택:",
                    ['전체', '위험', '주의', '양호']
                )
            
            # 필터 적용
            filtered_detail_data = data.copy()
            
            if selected_season != '전체':
                filtered_detail_data = filtered_detail_data[filtered_detail_data['시즌'] == selected_season]
            
            if selected_biz != '전체':
                filtered_detail_data = filtered_detail_data[filtered_detail_data['BIZ'] == selected_biz]
            
            if selected_status != '전체':
                status_map = {'위험': 'critical', '주의': 'warning', '양호': 'good'}
                filtered_detail_data = filtered_detail_data[filtered_detail_data['status'] == status_map[selected_status]]
            
            if len(filtered_detail_data) > 0:
                # 상세 분석 데이터 생성 (매장명 제외)
                detailed_analysis = []
                for _, row in filtered_detail_data.iterrows():
                    status_korean = {'critical': '위험', 'warning': '주의', 'good': '양호'}
                    detailed_analysis.append({
                        '시즌': row['시즌'],
                        'BIZ': row['BIZ'],
                        '상품코드': row['상품코드'],
                        '상품명': row['상품명'],
                        '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                        '현재_재고량': int(row['현재_재고량']),
                        '재고_수량': int(row['현재_재고량']),
                        '재고_커버리지_주': round(row['coverage_weeks'], 1),
                        '재고_상태': status_korean[row['status']],
                        '재고_금액': f"{int(row['재고_금액']):,}원"
                    })
                
                detailed_df = pd.DataFrame(detailed_analysis)
                st.write(f"**총 {len(detailed_df)}개 상품**")
                st.dataframe(detailed_df, use_container_width=True, hide_index=True)
            else:
                st.info("선택한 조건에 해당하는 데이터가 없습니다.")
            
            # 메뉴별 상세 분석 이동 (하단에 표시)
            st.subheader("📋 메뉴별 상세 분석 이동")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                if st.button("📊 전체 요약으로 이동"):
                    st.experimental_rerun()
            
            with col2:
                if st.button("🏢 BIZ별 분석으로 이동"):
                    st.experimental_rerun()
            
            with col3:
                if st.button("🌸 시즌별 분석으로 이동"):
                    st.experimental_rerun()
            
            with col4:
                if st.button("🏪 매장별 분석으로 이동"):
                    st.experimental_rerun()
        
        elif menu == "📄 종합 리포트":
            st.header("📄 종합 리포트")
            
            st.subheader("📊 주요 지표")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                total_inventory_value = data['재고_금액'].sum()
                st.metric("총 재고 금액", f"{total_inventory_value:,.0f}원")
            
            with col2:
                total_weekly_potential = (data['소비자가'] * data['avg_weekly_sales']).sum()
                st.metric("주간 예상 매출", f"{total_weekly_potential:,.0f}원")
            
            with col3:
                critical_ratio = len(data[data['status'] == 'critical']) / len(data) * 100
                st.metric("위험상품 비율", f"{critical_ratio:.1f}%")
            
            with col4:
                avg_coverage = data['coverage_weeks'].mean()
                st.metric("전체 평균 커버리지", f"{avg_coverage:.1f}주")
            
            st.subheader("📊 주요 지표 시각화")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**BIZ별 위험상품 수**")
                biz_critical = data[data['status'] == 'critical'].groupby('BIZ').size().reset_index(name='위험_상품수')
                
                # BIZ 순서 정렬 (AP, FW, EQ 순)
                biz_order = ['AP', 'FW', 'EQ']
                if len(set(biz_order) & set(biz_critical['BIZ'].unique())) > 0:
                    biz_critical['BIZ'] = pd.Categorical(biz_critical['BIZ'], categories=biz_order, ordered=True)
                    biz_critical = biz_critical.sort_values('BIZ')
                
                if len(biz_critical) > 0:
                    fig_critical = px.bar(
                        biz_critical,
                        x='BIZ',
                        y='위험_상품수',
                        title="BIZ별 위험상품 수",
                        color='위험_상품수',
                        color_continuous_scale='Reds',
                        text='위험_상품수'
                    )
                    fig_critical.update_traces(textposition='outside', textfont_size=16)
                    fig_critical.update_layout(yaxis=dict(range=[0, 3000]))
                    st.plotly_chart(fig_critical, use_container_width=True)
                else:
                    st.info("위험상품이 없습니다. ✅")
            
            with col2:
                st.write("**시즌별 재고 금액 분포**")
                season_inventory = data.groupby('시즌')['재고_금액'].sum().reset_index()
                season_inventory.columns = ['시즌', '재고_금액']
                season_inventory['재고_금액_억'] = season_inventory['재고_금액'] / 100000000
                season_inventory['재고_금액_표시'] = season_inventory['재고_금액_억'].apply(lambda x: f"{x:.1f}억")
                
                # BIZ별 위험상품 수와 동일한 막대그래프로 변경
                fig_inventory = px.bar(
                    season_inventory,
                    x='시즌',
                    y='재고_금액_억',
                    title="시즌별 재고 금액 분포",
                    color='재고_금액_억',
                    color_continuous_scale='Reds',
                    text='재고_금액_표시'
                )
                fig_inventory.update_traces(textposition='outside', textfont_size=16)
                fig_inventory.update_layout(yaxis=dict(range=[0, 800]))
                st.plotly_chart(fig_inventory, use_container_width=True)
            
            st.subheader("🎯 주요 개선 포인트")
            
            # 위험상품이 많은 BIZ별 SKU 수 표 - 재고 수량과 재고 금액 추가
            critical_by_biz = data[data['status'] == 'critical'].groupby('BIZ').size().sort_values(ascending=False)
            total_products = len(data)
            
            if len(critical_by_biz) > 0:
                st.write("**⚠️ 위험상품이 많은 BIZ별 SKU 수**")
                biz_critical_table = []
                
                biz_order = ['AP', 'FW', 'EQ']
                available_biz = [biz for biz in biz_order if biz in critical_by_biz.index]
                other_biz = [biz for biz in critical_by_biz.index if biz not in biz_order]
                all_biz = available_biz + other_biz
                
                for biz in all_biz:
                    count = critical_by_biz[biz]
                    percentage = (count / total_products) * 100
                    biz_inventory = data[data['BIZ'] == biz]['현재_재고량'].sum()
                    biz_amount = data[data['BIZ'] == biz]['재고_금액'].sum()
                    biz_critical_table.append({
                        'BIZ': biz,
                        '위험상품_SKU수': count,
                        '전체대비_비율': f"{percentage:.1f}%",
                        '재고_수량': int(biz_inventory),
                        '재고_금액': f"{int(biz_amount):,}원"
                    })
                
                # TOTAL 행 추가
                total_critical = len(data[data['status'] == 'critical'])
                total_percentage = (total_critical / total_products) * 100
                total_inventory = data['현재_재고량'].sum()
                total_amount = data['재고_금액'].sum()
                biz_critical_table.append({
                    'BIZ': 'TOTAL',
                    '위험상품_SKU수': total_critical,
                    '전체대비_비율': f"{total_percentage:.1f}%",
                    '재고_수량': int(total_inventory),
                    '재고_금액': f"{int(total_amount):,}원"
                })
                
                biz_critical_df = pd.DataFrame(biz_critical_table)
                
                # TOTAL 행 하이라이트
                def highlight_total_biz(row):
                    if row['BIZ'] == 'TOTAL':
                        return ['background-color: #000000; color: white; font-weight: bold;'] * len(row)
                    return [''] * len(row)
                
                styled_biz_critical = biz_critical_df.style.apply(highlight_total_biz, axis=1)
                st.dataframe(styled_biz_critical, use_container_width=True, hide_index=True)
            
            # 커버리지가 낮은 매장 (4주 미만) - 온라인 제외 표기 제외
            offline_data = data[data['매장명'] != '온라인']
            store_coverage = offline_data.groupby('매장명')['coverage_weeks'].mean().sort_values()
            low_coverage_stores = store_coverage[store_coverage < 4]
            overall_avg_coverage = offline_data['coverage_weeks'].mean()
            
            if len(low_coverage_stores) > 0:
                st.write("**📉 커버리지가 낮은 매장 (4주 미만)**")
                low_coverage_table = []
                for store, coverage in low_coverage_stores.head(10).items():
                    diff_from_avg = coverage - overall_avg_coverage
                    low_coverage_table.append({
                        '매장명': store,
                        '평균_커버리지': f"{coverage:.1f}주",
                        '전체평균_대비_차이': f"{diff_from_avg:+.1f}주"
                    })
                
                low_coverage_df = pd.DataFrame(low_coverage_table)
                st.dataframe(low_coverage_df, use_container_width=True, hide_index=True)
            else:
                st.write("**✅ 모든 매장의 커버리지가 4주 이상입니다.**")
            
            # 커버리지가 높은 매장 (상위 10개) - 온라인 제외 표기 제외와 ONLINE 제외한 전체 매장
            high_coverage_stores = store_coverage.tail(10)
            
            st.write("**📈 커버리지가 높은 매장 (상위 10개)**")
            high_coverage_table = []
            for store, coverage in high_coverage_stores.items():
                diff_from_avg = coverage - overall_avg_coverage
                high_coverage_table.append({
                    '매장명': store,
                    '평균_커버리지': f"{coverage:.1f}주",
                    '전체평균_대비_차이': f"{diff_from_avg:+.1f}주"
                })
            
            high_coverage_df = pd.DataFrame(high_coverage_table)
            st.dataframe(high_coverage_df, use_container_width=True, hide_index=True)
            
            # AP 위험 높은 상품 10개 - 재고 수량, 재고 금액, 보유 매장 수 표기 추가
            ap_data = data[data['BIZ'] == 'AP']
            if len(ap_data) > 0:
                st.write("**🚨 AP 위험 높은 상품 10개**")
                ap_critical = ap_data[ap_data['status'] == 'critical'].nlargest(10, 'avg_weekly_sales')
                if len(ap_critical) > 0:
                    ap_critical_table = []
                    for _, row in ap_critical.iterrows():
                        store_count = len(data[(data['상품코드'] == row['상품코드']) & (data['BIZ'] == 'AP')]['매장명'].unique())
                        ap_critical_table.append({
                            '상품코드': row['상품코드'],
                            '상품명': row['상품명'],
                            '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                            '현재_재고량': int(row['현재_재고량']),
                            '재고_커버리지_주': round(row['coverage_weeks'], 1),
                            '재고_수량': int(row['현재_재고량']),
                            '재고_금액': f"{int(row['재고_금액']):,}원",
                            '보유_매장수': store_count
                        })
                    ap_critical_df = pd.DataFrame(ap_critical_table)
                    st.dataframe(ap_critical_df, use_container_width=True, hide_index=True)
                else:
                    st.info("AP BIZ에 위험상품이 없습니다.")
            
            # AP 양호 높은 상품 10개 - 재고 수량, 재고 금액, 보유 매장 수 표기 추가
            if len(ap_data) > 0:
                st.write("**✅ AP 양호 높은 상품 10개**")
                ap_good = ap_data[ap_data['status'] == 'good'].nlargest(10, 'avg_weekly_sales')
                if len(ap_good) > 0:
                    ap_good_table = []
                    for _, row in ap_good.iterrows():
                        store_count = len(data[(data['상품코드'] == row['상품코드']) & (data['BIZ'] == 'AP')]['매장명'].unique())
                        ap_good_table.append({
                            '상품코드': row['상품코드'],
                            '상품명': row['상품명'],
                            '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                            '현재_재고량': int(row['현재_재고량']),
                            '재고_커버리지_주': round(row['coverage_weeks'], 1),
                            '재고_수량': int(row['현재_재고량']),
                            '재고_금액': f"{int(row['재고_금액']):,}원",
                            '보유_매장수': store_count
                        })
                    ap_good_df = pd.DataFrame(ap_good_table)
                    st.dataframe(ap_good_df, use_container_width=True, hide_index=True)
                else:
                    st.info("AP BIZ에 양호상품이 없습니다.")
            
            # FW 위험 높은 상품 10개 - 재고 수량, 재고 금액, 보유 매장 수 표기 추가
            fw_data = data[data['BIZ'] == 'FW']
            if len(fw_data) > 0:
                st.write("**🚨 FW 위험 높은 상품 10개**")
                fw_critical = fw_data[fw_data['status'] == 'critical'].nlargest(10, 'avg_weekly_sales')
                if len(fw_critical) > 0:
                    fw_critical_table = []
                    for _, row in fw_critical.iterrows():
                        store_count = len(data[(data['상품코드'] == row['상품코드']) & (data['BIZ'] == 'FW')]['매장명'].unique())
                        fw_critical_table.append({
                            '상품코드': row['상품코드'],
                            '상품명': row['상품명'],
                            '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                            '현재_재고량': int(row['현재_재고량']),
                            '재고_커버리지_주': round(row['coverage_weeks'], 1),
                            '재고_수량': int(row['현재_재고량']),
                            '재고_금액': f"{int(row['재고_금액']):,}원",
                            '보유_매장수': store_count
                        })
                    fw_critical_df = pd.DataFrame(fw_critical_table)
                    st.dataframe(fw_critical_df, use_container_width=True, hide_index=True)
                else:
                    st.info("FW BIZ에 위험상품이 없습니다.")
            
            # FW 양호 높은 상품 10개 - 재고 수량, 재고 금액, 보유 매장 수 표기 추가
            if len(fw_data) > 0:
                st.write("**✅ FW 양호 높은 상품 10개**")
                fw_good = fw_data[fw_data['status'] == 'good'].nlargest(10, 'avg_weekly_sales')
                if len(fw_good) > 0:
                    fw_good_table = []
                    for _, row in fw_good.iterrows():
                        store_count = len(data[(data['상품코드'] == row['상품코드']) & (data['BIZ'] == 'FW')]['매장명'].unique())
                        fw_good_table.append({
                            '상품코드': row['상품코드'],
                            '상품명': row['상품명'],
                            '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                            '현재_재고량': int(row['현재_재고량']),
                            '재고_커버리지_주': round(row['coverage_weeks'], 1),
                            '재고_수량': int(row['현재_재고량']),
                            '재고_금액': f"{int(row['재고_금액']):,}원",
                            '보유_매장수': store_count
                        })
                    fw_good_df = pd.DataFrame(fw_good_table)
                    st.dataframe(fw_good_df, use_container_width=True, hide_index=True)
                else:
                    st.info("FW BIZ에 양호상품이 없습니다.")
            
            # AP 커버리지 높은 상품 10개 - 재고 수량, 재고 금액, 보유 매장 수 표기 추가
            if len(ap_data) > 0:
                st.write("**📈 AP 커버리지 높은 상품 10개**")
                ap_high_coverage = ap_data[ap_data['coverage_weeks'] < 999].nlargest(10, 'coverage_weeks')
                if len(ap_high_coverage) > 0:
                    ap_high_coverage_table = []
                    for _, row in ap_high_coverage.iterrows():
                        store_count = len(data[(data['상품코드'] == row['상품코드']) & (data['BIZ'] == 'AP')]['매장명'].unique())
                        ap_high_coverage_table.append({
                            '상품코드': row['상품코드'],
                            '상품명': row['상품명'],
                            '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                            '현재_재고량': int(row['현재_재고량']),
                            '재고_커버리지_주': round(row['coverage_weeks'], 1),
                            '재고_수량': int(row['현재_재고량']),
                            '재고_금액': f"{int(row['재고_금액']):,}원",
                            '보유_매장수': store_count
                        })
                    ap_high_coverage_df = pd.DataFrame(ap_high_coverage_table)
                    st.dataframe(ap_high_coverage_df, use_container_width=True, hide_index=True)
            
            # AP 커버리지 낮은 상품 10개 - 재고 수량, 재고 금액, 보유 매장 수 표기 추가
            if len(ap_data) > 0:
                st.write("**📉 AP 커버리지 낮은 상품 10개**")
                ap_low_coverage = ap_data.nsmallest(10, 'coverage_weeks')
                if len(ap_low_coverage) > 0:
                    ap_low_coverage_table = []
                    for _, row in ap_low_coverage.iterrows():
                        store_count = len(data[(data['상품코드'] == row['상품코드']) & (data['BIZ'] == 'AP')]['매장명'].unique())
                        ap_low_coverage_table.append({
                            '상품코드': row['상품코드'],
                            '상품명': row['상품명'],
                            '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                            '현재_재고량': int(row['현재_재고량']),
                            '재고_커버리지_주': round(row['coverage_weeks'], 1),
                            '재고_수량': int(row['현재_재고량']),
                            '재고_금액': f"{int(row['재고_금액']):,}원",
                            '보유_매장수': store_count
                        })
                    ap_low_coverage_df = pd.DataFrame(ap_low_coverage_table)
                    st.dataframe(ap_low_coverage_df, use_container_width=True, hide_index=True)
            
            # FW 커버리지 높은 상품 10개 - 재고 수량, 재고 금액, 보유 매장 수 표기 추가
            if len(fw_data) > 0:
                st.write("**📈 FW 커버리지 높은 상품 10개**")
                fw_high_coverage = fw_data[fw_data['coverage_weeks'] < 999].nlargest(10, 'coverage_weeks')
                if len(fw_high_coverage) > 0:
                    fw_high_coverage_table = []
                    for _, row in fw_high_coverage.iterrows():
                        store_count = len(data[(data['상품코드'] == row['상품코드']) & (data['BIZ'] == 'FW')]['매장명'].unique())
                        fw_high_coverage_table.append({
                            '상품코드': row['상품코드'],
                            '상품명': row['상품명'],
                            '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                            '현재_재고량': int(row['현재_재고량']),
                            '재고_커버리지_주': round(row['coverage_weeks'], 1),
                            '재고_수량': int(row['현재_재고량']),
                            '재고_금액': f"{int(row['재고_금액']):,}원",
                            '보유_매장수': store_count
                        })
                    fw_high_coverage_df = pd.DataFrame(fw_high_coverage_table)
                    st.dataframe(fw_high_coverage_df, use_container_width=True, hide_index=True)
            
            # FW 커버리지 낮은 상품 10개 - 재고 수량, 재고 금액, 보유 매장 수 표기 추가
            if len(fw_data) > 0:
                st.write("**📉 FW 커버리지 낮은 상품 10개**")
                fw_low_coverage = fw_data.nsmallest(10, 'coverage_weeks')
                if len(fw_low_coverage) > 0:
                    fw_low_coverage_table = []
                    for _, row in fw_low_coverage.iterrows():
                        store_count = len(data[(data['상품코드'] == row['상품코드']) & (data['BIZ'] == 'FW')]['매장명'].unique())
                        fw_low_coverage_table.append({
                            '상품코드': row['상품코드'],
                            '상품명': row['상품명'],
                            '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                            '현재_재고량': int(row['현재_재고량']),
                            '재고_커버리지_주': round(row['coverage_weeks'], 1),
                            '재고_수량': int(row['현재_재고량']),
                            '재고_금액': f"{int(row['재고_금액']):,}원",
                            '보유_매장수': store_count
                        })
                    fw_low_coverage_df = pd.DataFrame(fw_low_coverage_table)
                    st.dataframe(fw_low_coverage_df, use_container_width=True, hide_index=True)
        
        elif menu == "📧 이메일 발송":
            st.header("📧 리포트 이메일 발송")
            
            st.write("**📊 현재 데이터 요약:**")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("총 매장 수", len(data['매장명'].unique()))
            with col2:
                st.metric("총 상품 수", len(data['상품코드'].unique()))
            with col3:
                st.metric("위험상품 수", len(data[data['status'] == 'critical']))
            with col4:
                st.metric("평균 커버리지", f"{data['coverage_weeks'].mean():.1f}주")
            
            st.subheader("📧 이메일 설정")
            
            with st.form("email_form"):
                col1, col2 = st.columns(2)
                
                with col1:
                    sender_email = st.text_input("발신자 이메일 (Gmail)", placeholder="your_email@gmail.com")
                    sender_password = st.text_input("앱 비밀번호", type="password", help="Gmail 앱 비밀번호를 입력하세요")
                
                with col2:
                    recipient_email = st.text_input("수신자 이메일", placeholder="recipient@company.com")
                
                st.markdown("""
                **📋 Gmail 앱 비밀번호 설정 방법:**
                1. Gmail 계정 → 보안 설정
                2. 2단계 인증 활성화
                3. 앱 비밀번호 생성
                4. 생성된 16자리 비밀번호 사용
                5. 앱비밀번호: dgnh kzzv fwyp lnbn
                """)
                
                submitted = st.form_submit_button("📧 리포트 발송", type="primary")
                
                if submitted:
                    if not all([sender_email, sender_password, recipient_email]):
                        st.error("❌ 모든 필드를 입력해주세요.")
                    else:
                        with st.spinner("이메일 발송 중..."):
                            try:
                                success, message = send_email_report(data, recipient_email, sender_email, sender_password)
                                
                                if success:
                                    st.success(f"✅ {message}")
                                    st.balloons()
                                else:
                                    st.error(f"❌ {message}")
                                    
                            except Exception as e:
                                st.error(f"❌ 이메일 발송 실패: {str(e)}")
            
            st.subheader("📥 빠른 다운로드")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                # 전체 분석 Excel 다운로드
                try:
                    excel_data = convert_df_to_excel(data, '전체분석')
                    st.download_button(
                        label="📊 전체 분석 다운로드 (Excel)",
                        data=excel_data,
                        file_name=f"전체_분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except:
                    csv_data = data.to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="📊 전체 분석 다운로드 (CSV)",
                        data=csv_data,
                        file_name=f"전체_분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
                    st.info("💡 Excel 다운로드를 위해 'pip install openpyxl' 실행하세요")
            
            with col2:
                # 상품코드별 분석 Excel 다운로드
                product_analysis = []
                for _, row in data.iterrows():
                    status_korean = {'critical': '위험', 'warning': '주의', 'good': '양호'}
                    product_analysis.append({
                        '시즌': row['시즌'],
                        'BIZ': row['BIZ'],
                        '상품코드': row['상품코드'],
                        '상품명': row['상품명'],
                        '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                        '현재_재고량': int(row['현재_재고량']),
                        '재고_커버리지_주': round(row['coverage_weeks'], 1),
                        '재고_상태': status_korean[row['status']],
                        '재고_금액': int(row['재고_금액'])
                    })
                
                product_df = pd.DataFrame(product_analysis)
                
                try:
                    product_excel = convert_df_to_excel(product_df, '상품코드별분석')
                    st.download_button(
                        label="🏷️ 상품코드별 분석 다운로드 (Excel)",
                        data=product_excel,
                        file_name=f"상품코드별_분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except:
                    product_csv = product_df.to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="🏷️ 상품코드별 분석 다운로드 (CSV)",
                        data=product_csv,
                        file_name=f"상품코드별_분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
            
            with col3:
                # 매장별 상세 데이터 Excel 다운로드
                detailed_analysis = []
                for _, row in data.iterrows():
                    status_korean = {'critical': '위험', 'warning': '주의', 'good': '양호'}
                    detailed_analysis.append({
                        '매장명': row['매장명'],
                        '시즌': row['시즌'],
                        'BIZ': row['BIZ'],
                        '상품코드': row['상품코드'],
                        '상품명': row['상품명'],
                        '평균_주간_판매량': round(row['avg_weekly_sales'], 1),
                        '현재_재고량': int(row['현재_재고량']),
                        '재고_커버리지_주': round(row['coverage_weeks'], 1),
                        '재고_상태': status_korean[row['status']],
                        '재고_금액': int(row['재고_금액'])
                    })
                
                detailed_df = pd.DataFrame(detailed_analysis)
                
                try:
                    detailed_excel = convert_df_to_excel(detailed_df, '매장별상세데이터')
                    st.download_button(
                        label="🏪 매장별 상세 데이터 다운로드 (Excel)",
                        data=detailed_excel,
                        file_name=f"매장별_상세_데이터_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except:
                    detailed_csv = detailed_df.to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="🏪 매장별 상세 데이터 다운로드 (CSV)",
                        data=detailed_csv,
                        file_name=f"매장별_상세_데이터_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
            
            with col4:
                # HTML 리포트 다운로드
                html_report = create_html_report(data)
                st.download_button(
                    label="📄 HTML 리포트 다운로드",
                    data=html_report,
                    file_name=f"재고_커버리지_리포트_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
                    mime="text/html"
                )

else:
    st.info("📁 왼쪽 사이드바에서 Excel 파일을 업로드하여 시작하세요.")
    
    st.subheader("📋 필수 데이터 형식")
    st.markdown("""
    **필수 컬럼:**
    - `매장명`: 매장 이름
    - `상품명`: 상품 이름  
    - `상품코드`: 고유 상품 코드
    - `BIZ`: 사업부 구분
    - `시즌`: 시즌 구분
    - `소비자가`: 상품 가격
    - `1주차_판매량`: 1주차 판매 수량
    - `2주차_판매량`: 2주차 판매 수량
    - `3주차_판매량`: 3주차 판매 수량
    - `현재_재고량`: 현재 재고 수량
    - `재고_금액`: 재고 금액
    
    **📊 주요 분석 기능:**
    - 재고 커버리지 자동 계산 (현재 재고량 ÷ 3주 평균 판매량)
    - 매장/BIZ/시즌별 상세 분석 및 시각화
    - 위험상품 식별 및 알림 (2주 미만 재고)
    - 종합 리포트 생성 및 Excel 다운로드
    - 완전한 이메일 발송 기능 (Excel 첨부파일)
    
    **📈 개선된 기능:**
    - 모든 차트에 숫자 크게 표시
    - 온라인 매장 제외 옵션
    - TOTAL 행 색상 하이라이트
    - 트리맵 시각화로 커버리지 분포 표시
    - 전체요약과 종합리포트가 포함된 이메일 리포트
    """)