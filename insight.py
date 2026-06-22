"""
SQA 펌웨어 분석 대시보드 — 💡 인사이트 분석 페이지 (단독 실행 버전)

이 파일은 인사이트 분석 메뉴만 따로 추출한 독립 실행 가능 파일입니다.

실행 방법:
    streamlit run app_insight_only.py

필요 파일:
    - SQA_test_results_clean.csv  (분석 대상 데이터 — 같은 폴더에 있어야 함)
    - requirements.txt 에 명시된 라이브러리 설치 필요

A74072 조희주
서강대학교 AI·SW대학원
"""


import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import os
import unicodedata
from datetime import datetime
from io import BytesIO

# 통계 분석
from scipy import stats
from scipy.stats import chi2_contingency
from sklearn.linear_model import LinearRegression

# python-pptx 임포트 (PPT 생성용)
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


# ──────────────────────────────────────────────────────────────────────────
# 페이지 설정
# ──────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SQA 펌웨어 분석 — 인사이트 분석",
    page_icon="💡",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ──────────────────────────────────────────────────────────────────────────
# 한글 폰트 설정
# ──────────────────────────────────────────────────────────────────────────
def setup_korean_font():
    font_path = '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'
    if os.path.exists(font_path):
        fm.fontManager.addfont(font_path)
        return 'NanumGothic'
    return 'DejaVu Sans'

font_name = setup_korean_font()
plt.rcParams['font.family'] = font_name
plt.rcParams['axes.unicode_minus'] = False
sns.set_style("whitegrid")


# ──────────────────────────────────────────────────────────────────────────
# 상수
# ──────────────────────────────────────────────────────────────────────────


setup_korean_font()


# ──────────────────────────────────────────────────────────────────────────
# 도메인 매핑 dict + Action Item 자동 매칭
# ──────────────────────────────────────────────────────────────────────────
KEYWORD_DOMAIN_MAP = {
    'line broken':     {'area': '좌표 보간/샘플링 레이트',     'priority': 'HIGH'},
    'jitter':          {'area': '노이즈 필터/디바운싱',         'priority': 'MID'},
    'ghost touch':     {'area': '노이즈 임계값/그라운드',       'priority': 'HIGH'},
    'no touch':        {'area': '감도 보정/입력 임계값',         'priority': 'CRITICAL'},
    '2 point로 인식':  {'area': '멀티 입력 분리 알고리즘',         'priority': 'MID'},
    'edge 과밀착':     {'area': '경계영역 보정 (강하게 반응)',    'priority': 'MID'},
    'edge 미밀착':     {'area': '경계영역 보정 (약하게 반응)',    'priority': 'MID'},
    'touch delay':     {'area': '응답속도/인터럽트 처리',         'priority': 'CRITICAL'},
}

SEVERITY_MAP = {
    'no touch': 5, 'touch delay': 4, 'ghost touch': 4, 'line broken': 3,
    '2 point로 인식': 3, 'edge 과밀착': 2, 'edge 미밀착': 2, 'jitter': 2,
}

TEST_AREA_MAP = {
    'Wet': '방수·방습', 'Linearity': '좌표 보간', 'Filter': '노이즈 필터링',
    'A/P/L': '인식 정확도', 'Separation': '멀티 입력 분리',
    'Sensitivity': '입력 감도', 'Palm': 'Palm Rejection',
}

KEYWORDS_ALL = list(KEYWORD_DOMAIN_MAP.keys())

# ==============================================================================
# 💡 자동 인사이트 생성 함수 (분석가 시선)
# ==============================================================================
# 모든 인사이트는 4단계 구조:
# 1. What (수치/사실)
# 2. So What (비교/맥락)
# 3. Why (원인 추정)
# 4. Now What (권장 행동)
# ==============================================================================

# ==============================================================================
# 💼 ACTION PLAYBOOK (Fail_Type별 비즈니스/개발 액션 매핑)
# ==============================================================================
ACTION_PLAYBOOK = {
    'no touch': {
        'business': '핵심 기능 실패 → 출시 지연 가능성 ⚠️ HIGH. 고객사 클레임 직결.',
        'dev': [
            '⚡ 즉시: 입력 임계값(threshold) 설정 점검 및 캘리브레이션 데이터 검증',
            '📋 단주: PASS 비율 높은 모델의 임계값 설정과 비교 분석',
            '🔄 장주: 환경 조건별 임계값 자동 보정 로직 도입 검토',
        ],
        'comparison': '안정 빌드의 임계값 설정 코드와 diff 비교 · 정상 모델의 캘리브레이션 파라미터 대조',
    },
    'touch delay': {
        'business': '응답성 저하 → 사용자 경험 악화 → 고객사 클레임 가능성 ↑',
        'dev': [
            '⚡ 즉시: 응답속도 알고리즘 프로파일링 (병목 함수 식별)',
            '📋 단주: 안정 빌드와 인터럽트 처리 로직 코드 diff',
            '🔄 장주: 폴링 주기 / 인터럽트 우선순위 재설계 검토',
        ],
        'comparison': '응답속도 정상 빌드 vs 현재 빌드의 핸들러 코드 비교 · CPU 사용률 프로파일 대조',
    },
    'ghost touch': {
        'business': '오작동 발생 → 사용자 신뢰도 하락 · 안전 이슈 가능성',
        'dev': [
            '⚡ 즉시: 노이즈 필터링 임계값 강화',
            '📋 단주: EMI/EMC 환경에서 발생 패턴 로깅 분석',
            '🔄 장주: 노이즈 필터 알고리즘 (예: Kalman filter) 도입 검토',
        ],
        'comparison': '노이즈 환경 테스트 통과 빌드와 필터 파라미터 비교',
    },
    'line broken': {
        'business': '드로잉/제스처 기능 손상 → 그리기 앱 등 활용 모델에서 사용성 ↓',
        'dev': [
            '⚡ 즉시: 좌표 보간(interpolation) 알고리즘 검토',
            '📋 단주: 끊김 발생 좌표 패턴 분석 (속도/방향)',
            '🔄 장주: 적응형 샘플링 레이트 조정 로직 검토',
        ],
        'comparison': '안정 빌드의 좌표 보간 임계값 비교 · Drawing 정상 모델의 샘플링 레이트 대조',
    },
    '2 point로 인식': {
        'business': '멀티 입력 처리 오류 → 게임/생산성 앱 등 멀티 입력 활용 모델에서 사용성 ↓',
        'dev': [
            '⚡ 즉시: 멀티 입력 분리 알고리즘 검토',
            '📋 단주: 입력점 거리/면적 임계값 점검',
            '🔄 장주: 접점 클러스터링 알고리즘 재설계 검토',
        ],
        'comparison': '멀티 입력 정상 동작 빌드의 클러스터링 파라미터 비교',
    },
    'edge 과밀착': {
        'business': '엣지 영역 인식 오류 → 풀스크린 UI 모델에서 사용성 영향',
        'dev': [
            '⚡ 즉시: Palm Rejection 알고리즘 임계값 점검',
            '📋 단주: 엣지 영역 좌표 보정 로직 검토',
            '🔄 장주: 엣지 검출 알고리즘 재학습 검토',
        ],
        'comparison': '엣지 정상 모델의 Palm Rejection 파라미터 비교',
    },
    'edge 미밀착': {
        'business': '엣지 영역 입력 손실 → 풀스크린 UI 모델에서 핵심 영역 미인식',
        'dev': [
            '⚡ 즉시: 엣지 영역 감도 부스팅 파라미터 점검',
            '📋 단주: 베젤리스 모델의 엣지 캘리브레이션 검증',
            '🔄 장주: 엣지 전용 보정 알고리즘 도입 검토',
        ],
        'comparison': '엣지 정상 모델과 베젤 두께 / 감도 파라미터 비교',
    },
    'jitter': {
        'business': '입력 떨림 → 정밀 작업(서명, 그림) 시 품질 저하',
        'dev': [
            '⚡ 즉시: 좌표 스무딩(smoothing) 필터 임계값 강화',
            '📋 단주: 떨림 발생 좌표 분포 분석',
            '🔄 장주: 적응형 스무딩 알고리즘 도입 검토',
        ],
        'comparison': '정밀 입력 정상 모델의 스무딩 파라미터 비교',
    },
}


def get_action_for_keyword(keyword):
    """Fail_Type 키워드에 해당하는 액션 가져오기"""
    return ACTION_PLAYBOOK.get(keyword, {
        'business': '관련 결함 발생으로 해당 모듈의 안정성에 영향',
        'dev': [
            '⚡ 즉시: 결함 발생 시나리오 재현 및 로깅 분석',
            '📋 단주: 안정 빌드와 관련 모듈 코드 비교',
            '🔄 장주: 모듈 단위 리팩토링 검토',
        ],
        'comparison': '안정 빌드의 관련 모듈 파라미터와 비교 분석 권장',
    })





# ──────────────────────────────────────────────────────────────────────────
# 인사이트 박스 + 영상 필터 패널 (공용 헬퍼)
# ──────────────────────────────────────────────────────────────────────────
def render_insight_box(title, finding, action, severity="info", **kwargs):
    """간결한 인사이트 박스 (한 문장 발견 + 한 줄 액션)
    
    title: 인사이트 제목 (예: "빌드 추세 분석 — 안정화 진행 중")
    finding: 한 문장 발견 (수치 + 의미 + 결론을 한 문장에)
    action: 다음 액션 한 줄
    severity: info/warning/critical/success
    **kwargs: 호환성을 위한 추가 인자 (무시됨, 옛 코드 호환)
    """
    color_map = {
        "info": ("#3498db", "#EBF5FF"),
        "warning": ("#f39c12", "#FFF5E6"),
        "critical": ("#e74c3c", "#FFEBEE"),
        "success": ("#27ae60", "#E8F8F0"),
    }
    border_color, bg_color = color_map.get(severity, color_map["info"])
    
    # HTML (한 줄로 작성 - 코드블록 방지)
    html = f'<div style="background-color:{bg_color};border-left:4px solid {border_color};border-radius:8px;padding:1rem 1.25rem;margin:1rem 0;font-family:\'Pretendard Variable\', sans-serif;"><div style="font-weight:700;font-size:1rem;color:#1a1a1a;margin-bottom:0.6rem;">💡 {title}</div><div style="font-size:0.95rem;color:#2c3e50;line-height:1.55;margin-bottom:0.5rem;">{finding}</div><div style="background-color:white;padding:0.6rem 0.8rem;border-radius:6px;border:1px solid {border_color};font-size:0.9rem;color:{border_color};font-weight:600;">🎯 다음 액션: <span style="color:#2c3e50;font-weight:500;">{action}</span></div></div>'
    
    st.markdown(html, unsafe_allow_html=True)




def render_video_filter_panel(df, context, key_prefix):
    """인사이트 박스 아래 인라인 영상 필터 패널 (5개 필터 + 영상 목록)
    
    df: 전체 데이터프레임 (FAIL만 자동 필터링)
    context: 인사이트의 컨텍스트 (자동 적용 기본값)
        예: {'IC': 'GT9XS', 'Model': 'Yoga_Pro9', 'Keyword': 'touch delay'}
    key_prefix: Streamlit 위젯 키 충돌 방지용 (인사이트마다 다른 값)
    """
    # FAIL만 추출 + video_id 보장
    fail_df = df[df['Result'] == 'FAIL'].copy()
    if 'video_id' not in fail_df.columns:
        fail_df['video_id'] = ''
    
    # 컨텍스트에서 자동 적용할 기본값 추출
    default_ic = [context['IC']] if context.get('IC') else []
    default_model = [context['Model']] if context.get('Model') else []
    default_build = [context['Build_Num']] if context.get('Build_Num') else []
    default_test_item = [context['Test_Item']] if context.get('Test_Item') else []
    default_keyword = [context['Keyword']] if context.get('Keyword') else []
    
    # 컨텍스트 요약 메시지
    ctx_parts = []
    if context.get('IC'):
        ctx_parts.append(f"IC={context['IC']}")
    if context.get('Model'):
        ctx_parts.append(f"Model={context['Model']}")
    if context.get('Keyword'):
        ctx_parts.append(f"결함={context['Keyword']}")
    if context.get('Test_Item'):
        ctx_parts.append(f"시나리오={context['Test_Item']}")
    ctx_msg = " · ".join(ctx_parts) if ctx_parts else "전체 결함"
    
    # 컨텍스트 기준 1차 카운트
    ctx_filtered = fail_df.copy()
    for col, val in [('IC', context.get('IC')), ('Model', context.get('Model')),
                      ('Build_Num', context.get('Build_Num')),
                      ('Test_Item', context.get('Test_Item')),
                      ('Keyword', context.get('Keyword'))]:
        if val:
            ctx_filtered = ctx_filtered[ctx_filtered[col] == val]
    
    # Expander로 인라인 펼침
    with st.expander(f"📁 관련 영상 목록 보기 ({ctx_msg}) — 컨텍스트 {len(ctx_filtered)}건", expanded=False):
        st.caption("💡 인사이트 컨텍스트가 자동 적용됐어요. 필터를 추가/변경해서 영상을 좁힐 수 있어요.")
        
        # 5개 필터 (필터 분석 메뉴와 동일 UI)
        col1, col2, col3 = st.columns(3)
        with col1:
            f_ic = st.multiselect(
                "IC",
                options=sorted(fail_df['IC'].unique()),
                default=default_ic,
                key=f"{key_prefix}_ic",
            )
            f_test_item = st.multiselect(
                "Test_Item",
                options=sorted(fail_df['Test_Item'].unique()),
                default=default_test_item,
                key=f"{key_prefix}_testitem",
            )
        with col2:
            f_model = st.multiselect(
                "Model",
                options=sorted(fail_df['Model'].unique()),
                default=default_model,
                key=f"{key_prefix}_model",
            )
            f_keyword = st.multiselect(
                "Fail_Type",
                options=sorted([k for k in fail_df['Keyword'].unique() if k]),
                default=default_keyword,
                key=f"{key_prefix}_keyword",
            )
        with col3:
            f_build = st.multiselect(
                "Build_Num",
                options=sorted(fail_df['Build_Num'].unique(), reverse=True),
                default=default_build,
                key=f"{key_prefix}_build",
            )
        
        # 필터 적용
        filtered = fail_df.copy()
        if f_ic:
            filtered = filtered[filtered['IC'].isin(f_ic)]
        if f_model:
            filtered = filtered[filtered['Model'].isin(f_model)]
        if f_build:
            filtered = filtered[filtered['Build_Num'].isin(f_build)]
        if f_test_item:
            filtered = filtered[filtered['Test_Item'].isin(f_test_item)]
        if f_keyword:
            filtered = filtered[filtered['Keyword'].isin(f_keyword)]
        
        # 결과 수
        st.markdown(f"**🎬 영상 {len(filtered)}건**")
        
        if len(filtered) == 0:
            st.info("선택한 조건에 해당하는 영상이 없어요. 필터를 완화해보세요.")
            return
        
        # 정렬 (최신 빌드순 기본)
        filtered = filtered.sort_values('Build_Num', ascending=False)
        
        # 영상 목록을 표로 표시
        display = filtered[['Build_Num', 'IC', 'Model', 'Test_Item', 'Keyword', 'video_id']].copy()
        display.columns = ['빌드', 'IC', '모델', 'Test_Item', 'Fail_Type', 'video_id']
        display['영상'] = display['video_id'].apply(get_video_link)
        display = display.drop(columns=['video_id'])
        
        st.dataframe(
            display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "영상": st.column_config.LinkColumn(
                    "영상 보기",
                    display_text="🎬 영상 열기",
                ),
                "빌드": st.column_config.NumberColumn("빌드", format="%d"),
            },
        )





# ──────────────────────────────────────────────────────────────────────────
# 인사이트 분석 함수들
# ──────────────────────────────────────────────────────────────────────────
def insight_build_trend(df, fail_df):
    """빌드별 ∩ 곡선 인사이트 (한 문장 + 액션)"""
    build_summary = df.groupby('Build_Num').apply(
        lambda x: (x['Result'] == 'FAIL').sum() / len(x) * 100
    ).round(2)
    if len(build_summary) < 2:
        return
    
    peak_build = build_summary.idxmax()
    peak_rate = build_summary.max()
    latest_build = build_summary.index[-1]
    latest_rate = build_summary.iloc[-1]
    diff = peak_rate - latest_rate
    
    if diff > 5:
        severity = "success"
        finding = f"빌드 <b>{peak_build}</b> ({peak_rate:.1f}%) 정점 후 최신 빌드 <b>{latest_build}</b> ({latest_rate:.1f}%)까지 <b>-{diff:.1f}%p 안정화</b> → 후속 빌드 수정 효과 확인됨"
        action = f"빌드 {peak_build}의 변경 사항 분석해서 안정 패턴을 표준 가이드로 정리"
    elif latest_rate > peak_rate * 0.9:
        severity = "warning"
        finding = f"최신 빌드 <b>{latest_build}</b> ({latest_rate:.1f}%)이 정점({peak_rate:.1f}%) 수준에 근접 → <b>회귀 가능성</b> ⚠️"
        action = f"빌드 {latest_build}의 변경 사항 회귀 검증 + 다음 빌드 출시 보류 검토"
    else:
        severity = "info"
        finding = f"빌드별 Fail율이 {min(build_summary):.1f}%~{peak_rate:.1f}% 범위에서 <b>안정세 유지</b> (최신 {latest_rate:.1f}%)"
        action = f"현 안정 빌드({latest_build})를 안정성 베이스라인으로 설정하고 모니터링 유지"
    
    render_insight_box(
        title=f"빌드 추세 분석",
        finding=finding,
        action=action,
        severity=severity
    )


def insight_ic_failrate(df, fail_df, key_suffix=""):
    """IC별 Fail율 인사이트 (한 문장 + 액션)"""
    ic_summary = df.groupby('IC').apply(
        lambda x: (x['Result'] == 'FAIL').sum() / len(x) * 100
    ).round(2)
    if len(ic_summary) < 2:
        return
    
    max_ic = ic_summary.idxmax()
    max_rate = ic_summary.max()
    avg_rate = ic_summary.mean()
    ratio = max_rate / avg_rate
    
    ic_fails = fail_df[fail_df['IC'] == max_ic]
    if len(ic_fails) == 0 or 'Keyword' not in ic_fails.columns:
        return
    top_keyword_counts = ic_fails['Keyword'].value_counts()
    if len(top_keyword_counts) == 0:
        return
    top_keyword = top_keyword_counts.index[0]
    top_keyword_pct = top_keyword_counts.iloc[0] / len(ic_fails) * 100
    area = KEYWORD_DOMAIN_MAP.get(top_keyword, {}).get('area', '관련 모듈')
    
    severity = "critical" if ratio > 1.3 else "warning"
    
    finding = f"<b>{max_ic}</b> Fail율 <b>{max_rate:.1f}%</b> (평균 {ratio:.1f}배) + <b>{top_keyword}</b> 결함이 {top_keyword_pct:.0f}% 집중 → <b>{area}</b> 모듈이 병목"
    action = f"{max_ic} 펌웨어의 {area} 알고리즘 우선 검토 + 안정 IC와 코드 diff 비교"
    
    render_insight_box(
        title=f"IC별 결함 분석",
        finding=finding,
        action=action,
        severity=severity
    )
    
    # 영상 필터 패널 (해당 IC + 가장 많은 결함으로 자동 필터)
    render_video_filter_panel(
        df=df,
        context={'IC': max_ic, 'Keyword': top_keyword},
        key_prefix=f"video_ic_{max_ic}_{top_keyword}_{key_suffix}",
    )


def insight_failtype_impact(fail_df, key_suffix=""):
    """Fail_Type 영향 분석 — 위험도 가중 건수(risk score) 포함"""
    if 'Keyword' not in fail_df.columns:
        return
    fail_kw = fail_df[fail_df['Keyword'] != '']

    # 미분류 비중 경고
    unclassified = fail_df[fail_df['Keyword'] == '']
    unclassified_pct = len(unclassified) / len(fail_df) * 100 if len(fail_df) > 0 else 0
    if unclassified_pct >= 20:
        render_insight_box(
            title="미분류 결함 비중 경고",
            finding=f"Keyword가 비어있는 Fail 건이 전체의 <b>{unclassified_pct:.1f}%</b>({len(unclassified)}건) — "
                    f"원인 분류가 안 된 결함이 많아 인사이트 신뢰도가 낮아집니다",
            action="SQA 팀과 Keyword 분류 기준 재정립 후 미분류 건 소급 분류 진행",
            severity="warning"
        )

    if len(fail_kw) == 0:
        return

    kw_counts = fail_kw['Keyword'].value_counts()
    kw_models = fail_kw.groupby('Keyword')['Model'].nunique()

    # 위험도 가중 점수 (건수 × severity 가중치)
    risk_scores = {}
    for kw, cnt in kw_counts.items():
        weight = SEVERITY_MAP.get(kw, 1)
        risk_scores[kw] = cnt * weight
    top_risk_kw = max(risk_scores, key=risk_scores.get)
    top_risk_score = risk_scores[top_risk_kw]

    top_kw = kw_counts.index[0]
    top_count = kw_counts.iloc[0]
    top_models = kw_models[top_kw]

    wide_kw = kw_models.idxmax()
    wide_models = kw_models[wide_kw]

    area = KEYWORD_DOMAIN_MAP.get(top_kw, {}).get('area', '관련 모듈')

    if top_kw == wide_kw:
        finding = (
            f"<b>{top_kw}</b>가 빈도 1위({top_count}건) + 영향 범위 1위({top_models}개 모델) → "
            f"<b>{area}</b> 모듈의 광범위 구조적 이슈 · "
            f"위험도 가중 점수(건수×심각도) 최고: <b>{top_risk_kw}</b> ({top_risk_score}점)"
        )
        action = f"{area} 모듈 최우선 개선 + 영향 모델 {top_models}개 일괄 회귀 테스트"
    else:
        finding = (
            f"빈도 1위 <b>{top_kw}</b>({top_count}건, {top_models}개 모델) vs "
            f"범위 1위 <b>{wide_kw}</b>({wide_models}개 모델) → 두 결함의 원인이 다름 · "
            f"위험도 가중 최고: <b>{top_risk_kw}</b> ({top_risk_score}점)"
        )
        action = (
            f"{top_kw}는 빈도 집중 분석, {wide_kw}는 광범위 영향 분석으로 이원화 대응 · "
            f"위험도 가중 기준 {top_risk_kw} 우선 처리"
        )

    render_insight_box(
        title="결함 영향 분석 (위험도 가중 포함)",
        finding=finding,
        action=action,
        severity="warning"
    )
    
    # 영상 필터 패널 (위험도 가중 1위 결함으로 자동 필터)
    render_video_filter_panel(
        df=fail_df,
        context={'Keyword': top_risk_kw},
        key_prefix=f"video_failtype_{top_risk_kw}_{key_suffix}",
    )


def insight_chisquare(df, fail_df):
    """카이제곱 대조 분석 — 실제 계산값 동적 바인딩"""
    from scipy.stats import chi2_contingency

    def _chi2_calc(data, var1, var2):
        try:
            ct = pd.crosstab(data[var1], data[var2])
            if ct.size == 0 or min(ct.shape) < 2:
                return None, None
            chi2, p, dof, _ = chi2_contingency(ct)
            n = ct.sum().sum()
            v = np.sqrt(chi2 / (n * (min(ct.shape) - 1)))
            return round(v, 3), round(p, 4)
        except Exception:
            return None, None

    fail_kw = fail_df[fail_df['Keyword'] != '']

    # 4개 조합 계산
    v_ic_result,   p_ic_result   = _chi2_calc(df,      'IC',       'Result')
    v_ic_kw,       p_ic_kw       = _chi2_calc(fail_kw, 'IC',       'Keyword')
    v_cust_result, p_cust_result = _chi2_calc(df,      'Customer', 'Result')
    v_cust_kw,     p_cust_kw     = _chi2_calc(fail_kw, 'Customer', 'Keyword')

    # 합격률 관련 유의성 판단 (IC + 고객사 중 하나라도 유의하면)
    result_sig_parts = []
    if p_ic_result is not None:
        if p_ic_result < 0.05:
            result_sig_parts.append(f"IC (p={p_ic_result:.4f}, V={v_ic_result:.3f})")
        else:
            result_sig_parts.append(f"IC (p={p_ic_result:.4f} → 무관)")
    if p_cust_result is not None:
        if p_cust_result < 0.05:
            result_sig_parts.append(f"고객사 (p={p_cust_result:.4f}, V={v_cust_result:.3f})")
        else:
            result_sig_parts.append(f"고객사 (p={p_cust_result:.4f} → 무관)")

    # 결함 종류 관련 유의성 판단
    kw_sig_parts = []
    if p_ic_kw is not None:
        if p_ic_kw < 0.05:
            kw_sig_parts.append(f"IC (p={p_ic_kw:.4f}, V={v_ic_kw:.3f})")
    if p_cust_kw is not None:
        if p_cust_kw < 0.05:
            kw_sig_parts.append(f"고객사 (p={p_cust_kw:.4f}, V={v_cust_kw:.3f})")

    result_text = " · ".join(result_sig_parts) if result_sig_parts else "분석 불가"
    kw_text     = " · ".join(kw_sig_parts)     if kw_sig_parts     else "유의한 관계 없음"

    # severity: 결함 종류와 그룹 간 관계가 강할수록 critical
    max_v_kw = max(
        (v for v in [v_ic_kw, v_cust_kw] if v is not None),
        default=0
    )
    severity = "critical" if max_v_kw >= 0.5 else "warning" if max_v_kw >= 0.2 else "info"

    finding = (
        f"<b>합격률</b>: {result_text} — "
        f"<b>결함 종류</b>: {kw_text} → "
        f"그룹별로 <b>결함 패턴이 다르게 나타남</b> (합격률보다 결함 종류에 그룹 차이 집중)"
    )
    action = (
        "단순 합격률 KPI가 아닌 IC·고객사별 결함 패턴 카드 도입 → "
        "Cramér's V가 높은 조합의 결함 유형을 그룹별 맞춤 대응 전략으로 연결"
    )

    render_insight_box(
        title="합격률 vs 결함 종류 — 통계 대조 분석 (실측값 기반)",
        finding=finding,
        action=action,
        severity=severity
    )


def insight_monthly_trend(df, fail_df):
    """월별 추이 인사이트 (한 문장 + 액션)"""
    df_copy = df.copy()
    df_copy['Test_Date'] = pd.to_datetime(df_copy['Test_Date'], errors='coerce')
    df_copy = df_copy.dropna(subset=['Test_Date'])
    if len(df_copy) == 0:
        return
    df_copy['month'] = df_copy['Test_Date'].dt.to_period('M')
    monthly = df_copy.groupby('month').apply(
        lambda x: (x['Result'] == 'FAIL').sum() / len(x) * 100
    ).round(2)
    if len(monthly) < 2:
        return
    
    peak_month = monthly.idxmax()
    peak_rate = monthly.max()
    latest_month = monthly.index[-1]
    latest_rate = monthly.iloc[-1]
    avg_rate = monthly.mean()
    diff = latest_rate - avg_rate
    
    if diff < -3:
        severity = "success"
        finding = f"<b>{peak_month}</b> 정점({peak_rate:.1f}%) 이후 <b>{latest_month}</b> {latest_rate:.1f}%로 평균 대비 {diff:+.1f}%p <b>개선세</b>"
        action = f"{peak_month} 이후 적용된 펌웨어 수정 사항을 회사 표준 가이드로 정리"
    elif diff > 3:
        severity = "warning"
        finding = f"<b>{latest_month}</b> Fail율 {latest_rate:.1f}%, 평균({avg_rate:.1f}%) 대비 {diff:+.1f}%p <b>악화</b> ⚠️"
        action = f"{latest_month} 출시 빌드의 변경 사항 즉시 회귀 검증"
    else:
        severity = "info"
        finding = f"월별 Fail율이 {monthly.min():.1f}%~{peak_rate:.1f}% 범위에서 <b>안정 유지</b> (최근 {latest_month} {latest_rate:.1f}%)"
        action = "현 품질 수준을 회사 SQA 표준 베이스라인으로 확정 후 신규 기능 검토"
    
    render_insight_box(
        title=f"월별 Fail율 추이",
        finding=finding,
        action=action,
        severity=severity
    )


def insight_customer_analysis(df, fail_df):
    """고객사 분석 (한 문장 + 액션)"""
    cust_summary = df.groupby('Customer').apply(
        lambda x: (x['Result'] == 'FAIL').sum() / len(x) * 100
    ).round(2)
    if len(cust_summary) < 2:
        return
    
    max_cust = cust_summary.idxmax()
    max_rate = cust_summary.max()
    min_cust = cust_summary.idxmin()
    min_rate = cust_summary.min()
    spread = max_rate - min_rate
    
    max_fails = fail_df[fail_df['Customer'] == max_cust]
    if len(max_fails) > 0 and 'Keyword' in max_fails.columns:
        max_kw_series = max_fails['Keyword'].value_counts()
        if len(max_kw_series) > 0:
            max_kw = max_kw_series.index[0]
            area = KEYWORD_DOMAIN_MAP.get(max_kw, {}).get('area', '관련 모듈')
        else:
            max_kw = "다양한 결함"
            area = "다양한 모듈"
    else:
        max_kw = "다양한 결함"
        area = "다양한 모듈"
    
    severity = "warning" if spread > 5 else "info"
    
    finding = f"고객사별 Fail율 격차 <b>{spread:.1f}%p</b> ({min_cust} {min_rate:.1f}% ~ {max_cust} {max_rate:.1f}%) → {max_cust}의 주요 결함은 <b>{max_kw}</b>"
    action = f"{max_cust} 전용 {area} 영역 검증 시나리오 강화 + 모델 라인업 차이 분석"
    
    render_insight_box(
        title="고객사별 결함 패턴",
        finding=finding,
        action=action,
        severity=severity
    )


def insight_test_item_matrix(fail_df, key_suffix=""):
    """Test_Item × Fail_Type 매트릭스 (한 문장 + 액션)"""
    if 'Test_Item' not in fail_df.columns or 'Keyword' not in fail_df.columns:
        return
    fail_kw = fail_df[fail_df['Keyword'] != '']
    if len(fail_kw) == 0:
        return
    
    pair_counts = fail_kw.groupby(['Test_Item', 'Keyword']).size().sort_values(ascending=False)
    if len(pair_counts) == 0:
        return
    
    top_pair = pair_counts.index[0]
    top_count = pair_counts.iloc[0]
    test_item, keyword = top_pair
    area = KEYWORD_DOMAIN_MAP.get(keyword, {}).get('area', '관련 모듈')
    
    finding = f"<b>{test_item}</b> × <b>{keyword}</b> = <b>{top_count}건</b> 핫스팟 발견 → {area} 모듈의 약점이 해당 시나리오에서 자극됨"
    action = f"{test_item} 시나리오의 코드 경로 + {area} 모듈 우선 디버깅"
    
    render_insight_box(
        title="Test_Item × Fail_Type 핫스팟",
        finding=finding,
        action=action,
        severity="warning"
    )
    
    # 영상 필터 패널 (Test_Item × Keyword 자동 필터)
    render_video_filter_panel(
        df=fail_df,
        context={'Test_Item': test_item, 'Keyword': keyword},
        key_prefix=f"video_testitem_{test_item}_{keyword}_{key_suffix}",
    )





# ──────────────────────────────────────────────────────────────────────────
# 데이터 로드 + 세션 상태
# ──────────────────────────────────────────────────────────────────────────
def load_base_data():
    df = pd.read_csv(
        'SQA_test_results_clean.csv',
        parse_dates=['Test_Date'],
        dtype={'IC_Code': str, 'Model_Minor': str}
    )
    df['IC_Code'] = df['IC_Code'].str.zfill(2)
    df['Model_Minor'] = df['Model_Minor'].str.zfill(2)
    df['Keyword'] = df['Keyword'].fillna('')
    # Video_Link → video_id 자동 변환 (사이트 호환성)
    df = ensure_video_id(df)
    return df


def ensure_video_id(df):
    """
    Video_Link 컬럼에서 video_id를 자동 생성.
    원본 데이터와 업로드 데이터 모두에 적용되어 사이트가 일관된 방식으로
    Google Drive 링크를 생성할 수 있도록 함.
    """
    import hashlib
    
    if 'video_id' not in df.columns:
        df['video_id'] = ''
    
    needs_id = (
        (df['Result'] == 'FAIL') &
        ((df['video_id'].isna()) | (df['video_id'] == ''))
    )
    
    if needs_id.any():
        def make_id(row):
            link = row.get('Video_Link', '') if 'Video_Link' in row.index else ''
            if pd.notna(link) and link != '':
                fname = str(link).split('/')[-1].replace('.mp4', '')
                h = hashlib.md5(fname.encode()).hexdigest()
                return f"1{h[:12]}_demo_{h[12:20]}"
            seed = f"{row.get('No', '')}_{row.get('Model', '')}_{row.get('Keyword', '')}_{row.get('Build_Num', '')}"
            h = hashlib.md5(seed.encode()).hexdigest()
            return f"1{h[:12]}_demo_{h[12:20]}"
        
        df.loc[needs_id, 'video_id'] = df.loc[needs_id].apply(make_id, axis=1)
    
    return df


def get_video_link(video_id):
    if pd.isna(video_id) or video_id is None or video_id == '':
        return None
    return f"https://drive.google.com/file/d/{video_id}/view"


# ──────────────────────────────────────────────────────────────────────────
# 세션 상태 초기화
# ──────────────────────────────────────────────────────────────────────────
if 'uploaded_data' not in st.session_state:
    st.session_state.uploaded_data = []  # 누적된 업로드 데이터들
if 'data_version' not in st.session_state:
    st.session_state.data_version = 0  # 데이터 갱신 추적용


def get_current_df():
    """기본 데이터 + 업로드된 데이터 모두 합쳐 반환 (영상 ID 자동 보장)"""
    base_df = load_base_data()
    if st.session_state.uploaded_data:
        uploaded_with_id = [ensure_video_id(d.copy()) for d in st.session_state.uploaded_data]
        all_dfs = [base_df] + uploaded_with_id
        return pd.concat(all_dfs, ignore_index=True)
    return base_df


# ==============================================================================
# 🎯 메뉴 1: 필터 분석
# ==============================================================================



# ──────────────────────────────────────────────────────────────────────────
# 필터 분석 (Tab 1)
# ──────────────────────────────────────────────────────────────────────────
def render_filter_analysis(df, fail_df):
    plt.rcParams['font.family'] = font_name

    # 상단 KPI
    total_tests_all = len(df)
    total_fails_all = len(fail_df)
    fail_rate_all = total_fails_all / total_tests_all * 100

    st.markdown("### 📊 전체 데이터 현황")
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown("##### 📋 전체 테스트")
        st.markdown(f"### :blue[**{total_tests_all:,}건**]")
    with k2:
        st.markdown("##### ❌ 전체 Fail")
        st.markdown(f"### :red[**{total_fails_all:,}건**]")
    with k3:
        st.markdown("##### 📈 전체 Fail율")
        st.markdown(f"### :orange[**{fail_rate_all:.1f}%**]")
    with k4:
        st.markdown("##### 💻 분석 모델")
        st.markdown(f"### :green[**{df['Model'].nunique()}개**]")

    st.markdown("---")

    # 필터 UI (다중 선택 가능)
    st.markdown("### 🎯 필터 선택")
    st.caption("각 필터에서 여러 값을 동시에 선택할 수 있습니다. (비어두면 전체)")

    row1_col1, row1_col2, row1_col3 = st.columns(3)
    with row1_col1:
        ic_filter = st.multiselect(
            "🔌 IC", 
            options=sorted(df['IC'].unique().tolist()),
            placeholder="전체 IC (비어두면 모두 선택)"
        )
    with row1_col2:
        model_filter = st.multiselect(
            "💻 Model",
            options=sorted(df['Model'].unique().tolist()),
            placeholder="전체 모델 (비어두면 모두 선택)"
        )
    with row1_col3:
        build_filter = st.multiselect(
            "🔄 Build",
            options=sorted(df['Build_Num'].unique().tolist()),
            placeholder="전체 빌드 (비어두면 모두 선택)"
        )
    
    row2_col1, row2_col2, row2_col3 = st.columns(3)
    with row2_col1:
        test_item_filter = st.multiselect(
            "🧪 Test_Item",
            options=sorted(df['Test_Item'].unique().tolist()),
            placeholder="전체 Test_Item (비어두면 모두 선택)"
        )
    with row2_col2:
        kw_filter = st.multiselect(
            "🐛 Fail_Type",
            options=KEYWORDS_ALL,
            placeholder="전체 Fail_Type (비어두면 모두 선택)"
        )
    with row2_col3:
        st.write("")

    # 필터 적용 (multiselect는 리스트)
    filtered_all = df.copy()
    filtered_fail = fail_df.copy()
    active_filters = []
    if ic_filter:
        filtered_all = filtered_all[filtered_all['IC'].isin(ic_filter)]
        filtered_fail = filtered_fail[filtered_fail['IC'].isin(ic_filter)]
        active_filters.append(f"IC={','.join(ic_filter)}")
    if model_filter:
        filtered_all = filtered_all[filtered_all['Model'].isin(model_filter)]
        filtered_fail = filtered_fail[filtered_fail['Model'].isin(model_filter)]
        if len(model_filter) <= 3:
            active_filters.append(f"Model={','.join(model_filter)}")
        else:
            active_filters.append(f"Model={len(model_filter)}개")
    if build_filter:
        filtered_all = filtered_all[filtered_all['Build_Num'].isin(build_filter)]
        filtered_fail = filtered_fail[filtered_fail['Build_Num'].isin(build_filter)]
        if len(build_filter) <= 3:
            active_filters.append(f"Build={','.join(map(str, build_filter))}")
        else:
            active_filters.append(f"Build={len(build_filter)}개")
    if test_item_filter:
        filtered_all = filtered_all[filtered_all['Test_Item'].isin(test_item_filter)]
        filtered_fail = filtered_fail[filtered_fail['Test_Item'].isin(test_item_filter)]
        if len(test_item_filter) <= 3:
            active_filters.append(f"Test_Item={','.join(test_item_filter)}")
        else:
            active_filters.append(f"Test_Item={len(test_item_filter)}개")
    if kw_filter:
        filtered_fail = filtered_fail[filtered_fail['Keyword'].isin(kw_filter)]
        if len(kw_filter) <= 3:
            active_filters.append(f"Fail_Type={','.join(kw_filter)}")
        else:
            active_filters.append(f"Fail_Type={len(kw_filter)}개")

    st.markdown("---")
    if active_filters:
        st.markdown(f"### 🎯 적용된 필터: `{' · '.join(active_filters)}`")
    else:
        st.markdown("### 🎯 전체 데이터 분석 (필터 미적용)")

    if len(filtered_fail) == 0:
        st.warning("⚠️ 선택한 조건에 해당하는 Fail 케이스가 없습니다.")
        return

    # 필터 결과 KPI
    st.markdown("### 📊 ① 필터 결과 요약")
    total_tests = len(filtered_all)
    total_fails = len(filtered_fail)
    fail_rate = (total_fails / total_tests * 100) if total_tests > 0 else 0
    n_models = filtered_fail['Model'].nunique()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown("##### 📋 전체 테스트")
        st.markdown(f"### :blue[**{total_tests:,}건**]")
    with k2:
        st.markdown("##### ❌ Fail 건수")
        st.markdown(f"### :red[**{total_fails:,}건**]")
    with k3:
        st.markdown("##### 📈 Fail율")
        st.markdown(f"### :orange[**{fail_rate:.1f}%**]")
    with k4:
        st.markdown("##### 💻 영향 모델")
        st.markdown(f"### :green[**{n_models}개**]")

    st.markdown("---")

    # Fail_Type 순위
    if not kw_filter:
        st.markdown("### 🏆 ② Fail_Type 순위 (TOP 5)")
        kw_rank = filtered_fail['Keyword'].value_counts().head(5).reset_index()
        kw_rank.columns = ['Fail_Type', '건수']
        kw_rank['비중'] = (kw_rank['건수'] / kw_rank['건수'].sum() * 100).round(1)
        kw_rank.insert(0, '순위', range(1, len(kw_rank) + 1))

        col_left, col_right = st.columns([1, 1])
        with col_left:
            fig, ax = plt.subplots(figsize=(6, 5))
            colors = plt.cm.Set3(np.linspace(0, 1, len(kw_rank)))
            ax.pie(kw_rank['건수'], labels=kw_rank['Fail_Type'],
                   autopct='%1.1f%%', wedgeprops=dict(width=0.4),
                   startangle=90, colors=colors,
                   textprops={'fontsize': 10})
            ax.set_title(f'Fail_Type 분포 (총 {len(filtered_fail)}건)',
                         fontsize=12, fontweight='bold')
            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
        with col_right:
            st.dataframe(kw_rank, use_container_width=True, hide_index=True)
        # 💡 인사이트 (차트 바로 아래)
        if len(filtered_fail[filtered_fail['Keyword'] != '']) > 0:
            insight_failtype_impact(filtered_fail, key_suffix="filter")
        st.markdown("---")

    # 빌드별 추이 (Fail율 % 기준 — 인사이트와 동일 기준)
    if not build_filter:
        st.markdown("### 📈 ③ 빌드별 Fail율 추이")
        st.caption("빌드별 테스트 물량이 다를 수 있으므로 절대 건수 대신 Fail율(%)로 표시합니다.")
        build_trend = filtered_all.groupby('Build_Num').apply(
            lambda x: pd.Series({
                'fail_rate': round((x['Result'] == 'FAIL').sum() / len(x) * 100, 1),
                'total': len(x),
                'fail_count': int((x['Result'] == 'FAIL').sum())
            })
        ).reset_index().sort_values('Build_Num')
        if len(build_trend) > 1:
            fig, ax = plt.subplots(figsize=(12, 4))
            ax.plot(build_trend['Build_Num'], build_trend['fail_rate'],
                    marker='o', markersize=8, linewidth=2, color='#1C7293')
            for _, row in build_trend.iterrows():
                ax.annotate(
                    f"{row['fail_rate']}%\n(n={int(row['total'])})",
                    (row['Build_Num'], row['fail_rate']),
                    textcoords="offset points", xytext=(0, 12),
                    ha='center', fontsize=9, fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                              edgecolor='#cccccc', alpha=0.85)
                )
            ax.set_xlabel('빌드 번호')
            ax.set_ylabel('Fail율 (%)')
            ax.set_ylim(0, max(build_trend['fail_rate']) * 1.35)
            ax.grid(True, alpha=0.3)
            ax.set_xticks(build_trend['Build_Num'])
            ax.set_xticklabels(build_trend['Build_Num'], rotation=45)
            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
            # 💡 인사이트 (차트와 동일 기준 — Fail율)
            if filtered_all['Build_Num'].nunique() >= 2:
                insight_build_trend(filtered_all, filtered_fail)
        st.markdown("---")

    # 모델별 순위 (Fail율 기준 정렬 + 건수 병기)
    if not model_filter:
        st.markdown("### 🔥 ④ 모델별 Fail 순위 (TOP 10) — Fail율 기준")
        st.caption("테스트 물량 차이를 보정하기 위해 Fail율(%) 기준으로 정렬합니다. 괄호 안은 절대 Fail 건수입니다.")
        model_stats = filtered_all.groupby('Model').apply(
            lambda x: pd.Series({
                'fail_count': int((x['Result'] == 'FAIL').sum()),
                'total': len(x),
                'fail_rate': round((x['Result'] == 'FAIL').sum() / len(x) * 100, 1)
            })
        ).reset_index()
        model_stats = model_stats[model_stats['fail_count'] > 0]
        model_rank = model_stats.sort_values('fail_rate', ascending=False).head(10)
        if len(model_rank) > 0:
            fig, ax = plt.subplots(figsize=(10, max(4, len(model_rank) * 0.4)))
            bars = ax.barh(model_rank['Model'], model_rank['fail_rate'],
                           color=plt.cm.YlOrRd(np.linspace(0.4, 0.9, len(model_rank))))
            for bar, rate, cnt in zip(bars, model_rank['fail_rate'], model_rank['fail_count']):
                ax.text(bar.get_width() + max(model_rank['fail_rate']) * 0.01,
                        bar.get_y() + bar.get_height()/2,
                        f'{rate}%  ({cnt}건)', va='center', fontsize=10, fontweight='bold')
            ax.set_xlabel('Fail율 (%)')
            ax.set_xlim(0, max(model_rank['fail_rate']) * 1.25)
            ax.invert_yaxis()
            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
            # 💡 인사이트 (IC별 결함 분석)
            if filtered_all['IC'].nunique() >= 2:
                insight_ic_failrate(filtered_all, filtered_fail, key_suffix="filter")
        st.markdown("---")
    if not test_item_filter:
        st.markdown("### 🧪 ⑤ Test_Item별 Fail 순위 (TOP 10)")
        item_rank = filtered_fail['Test_Item'].value_counts().head(10).reset_index()
        item_rank.columns = ['Test_Item', 'Fail 건수']
        if len(item_rank) > 0:
            fig, ax = plt.subplots(figsize=(10, max(4, len(item_rank) * 0.4)))
            bars = ax.barh(item_rank['Test_Item'], item_rank['Fail 건수'],
                           color=plt.cm.Blues(np.linspace(0.4, 0.9, len(item_rank))))
            for bar, cnt in zip(bars, item_rank['Fail 건수']):
                ax.text(bar.get_width() + max(item_rank['Fail 건수']) * 0.01,
                        bar.get_y() + bar.get_height()/2,
                        f'{cnt}건', va='center', fontsize=10, fontweight='bold')
            ax.set_xlabel('Fail 건수')
            ax.invert_yaxis()
            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
            # 💡 인사이트 (Test_Item × Fail_Type 매트릭스)
            if 'Keyword' in filtered_fail.columns and len(filtered_fail[filtered_fail['Keyword'] != '']) > 0:
                insight_test_item_matrix(filtered_fail, key_suffix="filter")
        st.markdown("---")
    st.markdown(f"### 📋 ⑥ 상세 Fail 케이스 ({len(filtered_fail):,}건)")
    st.caption("표의 영상 링크를 클릭하면 새 탭에서 결함 영상이 재생됩니다.")

    display_df = filtered_fail[[
        'No', 'IC', 'Model', 'Keyword', 'Build_Num',
        'Customer', 'Category', 'Test_Item', 'video_id'
    ]].copy()
    display_df['영상'] = display_df['video_id'].apply(get_video_link)
    display_df = display_df.drop(columns=['video_id'])
    # Fail_Type 건수 기준 정렬 (많은 것이 위로)
    kw_counts = filtered_fail['Keyword'].value_counts().to_dict()
    display_df['_sort'] = display_df['Keyword'].map(kw_counts).fillna(0)
    display_df = display_df.sort_values('_sort', ascending=False).drop(columns=['_sort'])

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "영상": st.column_config.LinkColumn("영상 보기", display_text="🎬 재생"),
            "Keyword": st.column_config.TextColumn("Fail_Type"),
        }
    )
    
    # ===== 카이제곱 대조 분석 (전체 데이터일 때만 표시 — 차트 외 별도 통계 분석) =====
    is_full = (not ic_filter and not model_filter and not build_filter 
               and not test_item_filter and not kw_filter)
    
    if is_full and len(filtered_all) > 0 and len(filtered_fail) > 0:
        st.markdown("---")
        st.markdown("### 📐 통계 검증 — 카이제곱 대조 분석")
        st.caption("필터를 비운 전체 데이터 기준 통계 검증입니다.")
        insight_chisquare(filtered_all, filtered_fail)


# ==============================================================================
# 📊 메뉴 2: 전체 인사이트 (생략 - 기존과 동일)
# ==============================================================================



# ──────────────────────────────────────────────────────────────────────────
# 전체 인사이트 (Tab 2 — 월별 추이 + 고객사 분석)
# ──────────────────────────────────────────────────────────────────────────
def render_full_insights(df, fail_df):
    plt.rcParams['font.family'] = font_name

    # 빌드별 ∩ 곡선
    st.markdown("### 1️⃣ 빌드별 ∩(역U) 안정화 곡선")
    build_summary = df.groupby('Build_Num').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    build_summary['fail_rate'] = (build_summary['fail'] / build_summary['total'] * 100).round(2)
    build_summary = build_summary.sort_values('Build_Num').reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(14, 5))
    ax.plot(build_summary['Build_Num'], build_summary['fail_rate'],
            marker='o', markersize=10, linewidth=2, color='#e74c3c')
    for i, row in build_summary.iterrows():
        if i == 0:
            offset_y, va = 15, 'bottom'
        else:
            prev_rate = build_summary.loc[i-1, 'fail_rate']
            offset_y = -20 if row['fail_rate'] < prev_rate else 15
            va = 'top' if row['fail_rate'] < prev_rate else 'bottom'
        ax.annotate(f"{row['fail_rate']}%",
                    (row['Build_Num'], row['fail_rate']),
                    textcoords="offset points", xytext=(0, offset_y),
                    ha='center', va=va, fontsize=11, fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                              edgecolor='#cccccc', alpha=0.9))
    ax.set_title('빌드별 전체 Fail율 추이', fontsize=14, fontweight='bold')
    ax.set_xlabel('빌드 번호')
    ax.set_ylabel('Fail율 (%)')
    ax.grid(True, alpha=0.3)
    ax.set_xticks(build_summary['Build_Num'])
    ax.set_xticklabels(build_summary['Build_Num'], rotation=45)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    
    # 💡 인사이트
    insight_build_trend(df, fail_df)
    
    st.markdown("---")

    # IC별 + Donut
    st.markdown("### 2️⃣ IC별 Fail율 + Fail_Type 분포")
    ic_summary = df.groupby('IC').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    ic_summary['fail_rate'] = (ic_summary['fail'] / ic_summary['total'] * 100).round(2)
    ic_summary = ic_summary.sort_values('fail_rate', ascending=False)

    fig1, ax1 = plt.subplots(figsize=(8, 4))
    bars = ax1.bar(ic_summary['IC'], ic_summary['fail_rate'],
                   color=['#e74c3c', '#f39c12', '#3498db'])
    ax1.set_title('IC별 Fail율', fontsize=14, fontweight='bold')
    ax1.set_ylabel('Fail율 (%)')
    ax1.set_ylim(0, max(ic_summary['fail_rate']) * 1.2)
    for bar, rate in zip(bars, ic_summary['fail_rate']):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                 f'{rate}%', ha='center', fontsize=12, fontweight='bold')
    plt.tight_layout()
    st.pyplot(fig1)
    plt.close(fig1)

    ic_keyword = fail_df.groupby(['IC', 'Keyword']).size().reset_index(name='count')
    colors_kw = plt.cm.Set3(np.linspace(0, 1, 8))
    color_dict = dict(zip(KEYWORDS_ALL, colors_kw))

    fig2, axes2 = plt.subplots(1, 3, figsize=(14, 5))
    fig2.suptitle('IC별 Fail_Type 분포', fontsize=14, fontweight='bold')
    for ax, ic in zip(axes2, ['G7500', 'GT1T0A', 'GT9XS']):
        ic_data = ic_keyword[ic_keyword['IC'] == ic].sort_values('count', ascending=False)
        total_cnt = ic_data['count'].sum()
        colors_this = [color_dict[kw] for kw in ic_data['Keyword']]
        ax.pie(ic_data['count'], labels=None,
               autopct=lambda pct: f'{pct:.1f}%' if pct >= 5 else '',
               wedgeprops=dict(width=0.4), startangle=90, colors=colors_this,
               textprops={'fontsize': 10, 'fontweight': 'bold'})
        ax.text(0, 0.1, ic, ha='center', va='center', fontsize=14, fontweight='bold')
        ax.text(0, -0.15, f'Fail {total_cnt}건', ha='center', va='center',
                fontsize=10, color='#666666')
    legend_patches = [plt.Rectangle((0,0),1,1, color=color_dict[kw]) for kw in KEYWORDS_ALL]
    fig2.legend(legend_patches, KEYWORDS_ALL, loc='lower center', ncol=4,
                fontsize=9, bbox_to_anchor=(0.5, -0.05))
    plt.tight_layout()
    st.pyplot(fig2)
    plt.close(fig2)
    
    # 💡 인사이트
    insight_ic_failrate(df, fail_df, key_suffix="stats")
    
    st.markdown("---")

    # 카테고리 히트맵
    st.markdown("### 3️⃣ 카테고리 × Fail_Type 히트맵")
    pivot_cat = fail_df.pivot_table(
        index='Category', columns='Keyword', values='No', aggfunc='count', fill_value=0)
    pivot_cat = pivot_cat[[k for k in KEYWORDS_ALL if k in pivot_cat.columns]]
    fig, ax = plt.subplots(figsize=(12, 6))
    sns.heatmap(pivot_cat, annot=True, fmt='d', cmap='YlOrRd',
                cbar_kws={'label': 'Fail 건수'}, ax=ax,
                linewidths=0.5, linecolor='white')
    ax.set_title('카테고리 × Fail_Type 히트맵', fontsize=14, fontweight='bold', pad=15)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    st.markdown("---")

    # 모델 히트맵
    st.markdown("### 4️⃣ 모델 × Fail_Type 히트맵 ⭐")
    pivot_model = fail_df.pivot_table(
        index='Model', columns='Keyword', values='No', aggfunc='count', fill_value=0)
    pivot_model = pivot_model.loc[pivot_model.sum(axis=1).sort_values(ascending=False).index]
    pivot_model = pivot_model[[k for k in KEYWORDS_ALL if k in pivot_model.columns]]
    fig, ax = plt.subplots(figsize=(12, max(8, len(pivot_model) * 0.3)))
    sns.heatmap(pivot_model, annot=True, fmt='d', cmap='YlOrRd',
                cbar_kws={'label': 'Fail 건수'}, ax=ax,
                linewidths=0.5, linecolor='white')
    ax.set_title('모델 × Fail_Type 히트맵', fontsize=14, fontweight='bold', pad=15)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    st.markdown("---")

    # 발생 횟수 × 영향 모델 수 버블 차트
    st.markdown("### 5️⃣ Fail_Type 영향 분석 (발생 횟수 × 영향 모델 수)")
    st.caption("오른쪽 위에 있을수록 자주 발생하고 많은 모델에 영향 → 우선 검토 대상")
    
    # 각 Fail_Type별 통계
    bubble_data = []
    for kw in KEYWORDS_ALL:
        kw_fails = fail_df[fail_df['Keyword'] == kw]
        if len(kw_fails) > 0:
            bubble_data.append({
                'keyword': kw,
                'count': len(kw_fails),
                'n_models': kw_fails['Model'].nunique()
            })
    
    if bubble_data:
        max_count = max(b['count'] for b in bubble_data)
        max_models = max(b['n_models'] for b in bubble_data)
        
        fig, ax = plt.subplots(figsize=(12, 7))
        colors = plt.cm.YlOrRd([b['count']/max_count for b in bubble_data])
        
        for i, b in enumerate(bubble_data):
            size = 300 + (b['count'] / max_count) * 2700
            ax.scatter(b['count'], b['n_models'], s=size, c=[colors[i]],
                       alpha=0.65, edgecolors='black', linewidths=1.5, zorder=3)
        
        # 라벨 (adjust_text 없이 수동으로 위치 분산)
        # y값 기준 정렬해서 가까운 것끼리 위치 어긋나게
        # 라벨 위치 분산 (각 버블마다 다른 방향)
        sorted_bubbles = sorted(bubble_data, key=lambda b: (b['n_models'], b['count']))
        # 8개 방향 분산
        directions = [
            (25, 25), (25, -25), (-25, 25), (-25, -25),
            (35, 5), (-35, 5), (5, 35), (5, -35)
        ]
        for i, b in enumerate(sorted_bubbles):
            offset_x, offset_y = directions[i % len(directions)]
            ha = 'left' if offset_x > 0 else 'right' 
            
            ax.annotate(
                f"{b['keyword']}\n({b['count']}건, {b['n_models']}개 모델)",
                (b['count'], b['n_models']),
                xytext=(offset_x, offset_y),
                textcoords='offset points',
                fontsize=9, fontweight='bold', ha=ha,
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                          edgecolor='#cccccc', alpha=0.9),
                zorder=4,
                arrowprops=dict(arrowstyle='-', color='#888888', lw=0.5, alpha=0.5)
            )
        
        # 평균선
        avg_count = sum(b['count'] for b in bubble_data) / len(bubble_data)
        avg_models = sum(b['n_models'] for b in bubble_data) / len(bubble_data)
        ax.axhline(y=avg_models, color='gray', linestyle='--', alpha=0.4,
                   label=f'평균 영향 모델 ({avg_models:.0f}개)')
        ax.axvline(x=avg_count, color='gray', linestyle='--', alpha=0.4,
                   label=f'평균 발생 횟수 ({avg_count:.0f}건)')
        
        ax.set_xlabel('발생 횟수 (Fail 건수)', fontsize=12, fontweight='bold')
        ax.set_ylabel('영향 모델 수', fontsize=12, fontweight='bold')
        ax.set_title('Fail_Type 영향 분석',
                     fontsize=14, fontweight='bold', pad=15)
        ax.set_xlim(0, max_count * 1.3)
        ax.set_ylim(0, max_models * 1.3)
        ax.grid(True, alpha=0.3)
        ax.legend(loc='lower right', fontsize=9)
        
        # 사분면 안내
        ax.text(max_count * 1.25, max_models * 1.25, '★ 자주 + 광범위\n우선 검토',
                fontsize=10, ha='right', va='top', color='#c0392b', fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.4', facecolor='#FFF5F5', edgecolor='#e74c3c'))
        
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
        
        # 표
        bubble_df = pd.DataFrame(bubble_data).sort_values('count', ascending=False).reset_index(drop=True)
        bubble_df.insert(0, '순위', range(1, len(bubble_df) + 1))
        bubble_df.columns = ['순위', 'Fail_Type', '발생 횟수', '영향 모델 수']
        st.dataframe(bubble_df, use_container_width=True, hide_index=True)
        
        # 💡 인사이트
        insight_failtype_impact(fail_df, key_suffix="stats")

    st.markdown("---")

    # ===== ⑥ 무엇이 결함을 결정하는가? (카이제곱 검정) =====
    st.markdown("### 6️⃣ 무엇이 결함을 결정하는가? (카이제곱 검정)")
    st.caption("IC/고객사가 '합격 여부'와 '결함 종류'에 각각 영향을 주는지 통계적으로 비교 분석합니다.")

    from scipy.stats import chi2_contingency

    def _run_chi2(data, var1, var2):
        try:
            contingency = pd.crosstab(data[var1], data[var2])
            if contingency.size == 0 or min(contingency.shape) < 2:
                return None, None
            chi2, p, dof, expected = chi2_contingency(contingency)
            n = contingency.sum().sum()
            cramers_v = np.sqrt(chi2 / (n * (min(contingency.shape) - 1)))
            return cramers_v, p
        except Exception:
            return None, None

    # 4개 검정
    fail_kw = fail_df[fail_df['Keyword'] != '']
    chi2_results = [
        ("IC ×\n합격률", *_run_chi2(df, 'IC', 'Result')),
        ("IC ×\n결함종류", *_run_chi2(fail_kw, 'IC', 'Keyword')),
        ("고객사 ×\n합격률", *_run_chi2(df, 'Customer', 'Result')),
        ("고객사 ×\n결함종류", *_run_chi2(fail_kw, 'Customer', 'Keyword')),
    ]
    chi2_results = [r for r in chi2_results if r[1] is not None]

    if chi2_results:
        labels = [r[0] for r in chi2_results]
        cramers = [r[1] for r in chi2_results]
        pvals = [r[2] for r in chi2_results]
        colors = ['#bdc3c7' if p >= 0.05 else '#e74c3c' for p in pvals]

        fig, ax = plt.subplots(figsize=(11, 6))
        bars = ax.bar(labels, cramers, color=colors, edgecolor='black',
                      linewidth=1.5, alpha=0.85)
        for bar, cv, p in zip(bars, cramers, pvals):
            sig = "유의함" if p < 0.05 else "무관"
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02,
                    f'{cv:.3f}\n({sig})', ha='center', fontsize=11, fontweight='bold')
        ax.axhline(y=0.1, color='gray', linestyle='--', alpha=0.5)
        ax.text(len(labels)-0.4, 0.11, '약한 관계 기준', fontsize=9, color='gray')
        ax.set_ylabel("Cramér's V (관계 강도)", fontsize=12, fontweight='bold')
        ax.set_title('합격률 vs 결함 종류 — 무엇이 IC/고객사와 관련 있는가',
                     fontsize=14, fontweight='bold')
        ax.set_ylim(0, max(cramers) * 1.25)
        # 범례
        from matplotlib.patches import Patch
        legend_elems = [
            Patch(facecolor='#e74c3c', edgecolor='black', label='유의함 (p<0.05)'),
            Patch(facecolor='#bdc3c7', edgecolor='black', label='무관 (p≥0.05)')
        ]
        ax.legend(handles=legend_elems, loc='upper left', fontsize=10)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

        # 💡 인사이트 (통일된 형식)
        insight_chisquare(df, fail_df)


# ==============================================================================
# 📈 메뉴: Fail율 예측
# ==============================================================================



# ──────────────────────────────────────────────────────────────────────────
# 전체 통계 (Tab 3 — 카이제곱 + Bubble + Test_Item 매트릭스 등)
# ──────────────────────────────────────────────────────────────────────────
def render_full_statistics(df, fail_df):
    """전체 데이터 통계 - 월별, Customer, Test_Item 매트릭스, 메타 정보"""
    plt.rcParams['font.family'] = font_name
    
    # ===== 종합 KPI (핵심 4개만) =====
    st.markdown("### 🎯 종합 KPI")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📋 전체 테스트", f"{len(df):,}건")
    with col2:
        fail_rate = len(fail_df) / len(df) * 100 if len(df) > 0 else 0
        st.metric("❌ Fail율", f"{fail_rate:.1f}%", f"{len(fail_df):,}건")
    with col3:
        st.metric("💻 모델 수", f"{df['Model'].nunique()}개")
    with col4:
        st.metric("🔄 빌드 수", f"{df['Build_Num'].nunique()}개")
    
    st.markdown("---")
    
    # ===== 1. 월별 테스트 추이 =====
    st.markdown("### 📅 1. 월별 테스트 추이")
    st.caption("시간 흐름에 따른 테스트량과 Fail율 변화")
    
    df_dated = df.copy()
    df_dated['Test_Date'] = pd.to_datetime(df_dated['Test_Date'], errors='coerce')
    df_dated = df_dated.dropna(subset=['Test_Date'])
    df_dated['year_month'] = df_dated['Test_Date'].dt.to_period('M').astype(str)
    
    monthly = df_dated.groupby('year_month').agg(
        total=('Result', 'count'),
        fail=('Result', lambda x: (x == 'FAIL').sum())
    ).reset_index()
    monthly['fail_rate'] = (monthly['fail'] / monthly['total'] * 100).round(1)
    
    fig, ax1 = plt.subplots(figsize=(14, 5))
    
    # 막대: 테스트 건수
    bars = ax1.bar(monthly['year_month'], monthly['total'],
                    color='#3498db', alpha=0.6, label='테스트 건수')
    ax1.set_xlabel('월', fontsize=11)
    ax1.set_ylabel('테스트 건수', fontsize=11, color='#3498db')
    ax1.tick_params(axis='y', labelcolor='#3498db')
    ax1.tick_params(axis='x', rotation=45)
    
    for bar, val in zip(bars, monthly['total']):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 20,
                 f'{val}', ha='center', fontsize=9, color='#3498db')
    
    # 라인: Fail율
    ax2 = ax1.twinx()
    ax2.plot(monthly['year_month'], monthly['fail_rate'],
             marker='o', markersize=8, linewidth=2, color='#e74c3c', label='Fail율 (%)')
    ax2.set_ylabel('Fail율 (%)', fontsize=11, color='#e74c3c')
    ax2.tick_params(axis='y', labelcolor='#e74c3c')
    ax2.set_ylim(0, max(monthly['fail_rate']) * 1.3)
    
    for i, val in enumerate(monthly['fail_rate']):
        ax2.text(i, val + 1, f'{val}%', ha='center', fontsize=9, color='#e74c3c', fontweight='bold')
    
    ax1.set_title('월별 테스트 건수 + Fail율 추이', fontsize=14, fontweight='bold')
    ax1.grid(True, alpha=0.3)
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)
    
    # 표
    monthly_display = monthly.copy()
    monthly_display.columns = ['월', '전체', 'Fail', 'Fail율(%)']
    st.dataframe(monthly_display, use_container_width=True, hide_index=True)
    
    # 💡 인사이트
    insight_monthly_trend(df, fail_df)
    
    st.markdown("---")
    
    # ===== 2. Customer (고객사) 분석 =====
    st.markdown("### 🏢 2. 고객사별 분석")
    st.caption("어떤 고객사가 어떤 IC를 어떻게 테스트하는지")
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("##### 고객사별 테스트 건수 및 Fail율")
        cust = df.groupby('Customer').agg(
            total=('Result', 'count'),
            fail=('Result', lambda x: (x == 'FAIL').sum())
        ).reset_index()
        cust['fail_rate'] = (cust['fail'] / cust['total'] * 100).round(1)
        cust = cust.sort_values('total', ascending=False)
        
        fig, ax = plt.subplots(figsize=(8, 5))
        bars = ax.barh(cust['Customer'], cust['total'],
                       color=plt.cm.Set2(np.linspace(0, 1, len(cust))))
        for bar, t, fr in zip(bars, cust['total'], cust['fail_rate']):
            ax.text(bar.get_width() + max(cust['total']) * 0.01,
                    bar.get_y() + bar.get_height()/2,
                    f'{t}건 ({fr}%)', va='center', fontsize=9, fontweight='bold')
        ax.set_xlabel('테스트 건수')
        ax.invert_yaxis()
        ax.set_title('고객사별 테스트 분포', fontsize=12, fontweight='bold')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
    
    with col_right:
        st.markdown("##### IC × 고객사 매트릭스")
        ic_cust = df.groupby(['IC', 'Customer']).size().unstack(fill_value=0)
        
        fig, ax = plt.subplots(figsize=(8, 5))
        sns.heatmap(ic_cust, annot=True, fmt='d', cmap='Blues',
                    cbar_kws={'label': '테스트 건수'}, ax=ax,
                    linewidths=0.5, linecolor='white')
        ax.set_title('IC × 고객사 분포', fontsize=12, fontweight='bold')
        ax.set_xlabel('Customer')
        ax.set_ylabel('IC')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
    
    # 💡 인사이트
    insight_customer_analysis(df, fail_df)
    
    st.markdown("---")
    
    # ===== 3. Test_Item × Fail_Type 매트릭스 =====
    st.markdown("### 🧪 3. Test_Item × Fail_Type 매트릭스")
    st.caption("어떤 테스트 항목에서 어떤 결함이 자주 발생하는지")
    
    if len(fail_df) > 0:
        # Top 15 Test_Item × Fail_Type
        top_items = fail_df['Test_Item'].value_counts().head(15).index.tolist()
        filtered = fail_df[fail_df['Test_Item'].isin(top_items)]
        
        item_kw = filtered.pivot_table(
            index='Test_Item', columns='Keyword', values='No',
            aggfunc='count', fill_value=0
        )
        item_kw = item_kw.loc[top_items]  # 순서 유지
        item_kw = item_kw[[k for k in KEYWORDS_ALL if k in item_kw.columns]]
        
        fig, ax = plt.subplots(figsize=(14, max(6, len(item_kw) * 0.4)))
        sns.heatmap(item_kw, annot=True, fmt='d', cmap='YlOrRd',
                    cbar_kws={'label': 'Fail 건수'}, ax=ax,
                    linewidths=0.5, linecolor='white')
        ax.set_title('Test_Item (TOP 15) × Fail_Type 매트릭스',
                     fontsize=14, fontweight='bold', pad=15)
        ax.set_xlabel('Fail_Type')
        ax.set_ylabel('Test_Item')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
        
        # 💡 인사이트 (단순 안내 → 의사결정용)
        insight_test_item_matrix(fail_df, key_suffix="testitem")
    
    st.markdown("---")
    
    # ===== 4. 데이터 메타 정보 =====
    st.markdown("### 📋 4. 데이터 메타 정보")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("##### 📅 수집 기간")
        if len(df_dated) > 0:
            min_date = df_dated['Test_Date'].min()
            max_date = df_dated['Test_Date'].max()
            duration = (max_date - min_date).days
            st.markdown(f"""
            - **시작**: {min_date.strftime('%Y-%m-%d')}
            - **종료**: {max_date.strftime('%Y-%m-%d')}
            - **기간**: 약 {duration}일 ({duration//30}개월)
            - **수집 월**: {df_dated['year_month'].nunique()}개월
            """)
    
    with col2:
        st.markdown("##### 🔌 IC 종류")
        for ic in sorted(df['IC'].unique()):
            ic_data = df[df['IC'] == ic]
            ic_fail = ic_data[ic_data['Result'] == 'FAIL']
            top_kw = ic_fail['Keyword'].value_counts().head(1)
            top_text = f"{top_kw.index[0]} ({top_kw.iloc[0]}건)" if len(top_kw) > 0 else "Fail 없음"
            st.markdown(f"- **{ic}**: {len(ic_data):,}건 · Top Fail: {top_text}")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### 📊 결과 분포")
        result_counts = df['Result'].value_counts()
        for r, c in result_counts.items():
            pct = c / len(df) * 100
            st.markdown(f"- **{r}**: {c:,}건 ({pct:.1f}%)")
    
    with col2:
        st.markdown("##### 🐛 주요 Fail_Type (TOP 5)")
        top_kw = fail_df['Keyword'].value_counts().head(5)
        for i, (kw, c) in enumerate(top_kw.items(), 1):
            st.markdown(f"- **{i}위 {kw}**: {c}건")




# ==============================================================================
# 📁 메뉴 4: 데이터 업로드 (선택 삭제 + 팝업 UX)
# ==============================================================================



# ──────────────────────────────────────────────────────────────────────────
# 헬퍼: 메뉴별 도움말 표시
# ──────────────────────────────────────────────────────────────────────────
def show_menu_help(title, description, tips):
    """메뉴 상단에 사용법 안내 박스 (HTML 들여쓰기 없이 한 줄로)"""
    if show_help:
        tips_html = "".join([f'<li style="margin-bottom:0.35rem;">{tip}</li>' for tip in tips])
        html = f'<div style="background-color:#f5f1e8;border:1px solid #d4d1c4;border-left:4px solid #cc785c;border-radius:10px;padding:1rem 1.25rem;margin-bottom:1rem;font-family:\'Pretendard Variable\', sans-serif;"><div style="font-weight:700;color:#1a1a1a;margin-bottom:0.5rem;font-size:0.95rem;">💡 사용법 안내</div><div style="color:#4a4a4a;font-size:0.875rem;margin-bottom:0.75rem;line-height:1.55;">{description}</div><ul style="color:#4a4a4a;font-size:0.875rem;line-height:1.55;margin:0;padding-left:1.25rem;">{tips_html}</ul></div>'
        st.markdown(html, unsafe_allow_html=True)





# ──────────────────────────────────────────────────────────────────────────
# 메인 화면 — 인사이트 분석 페이지
# ──────────────────────────────────────────────────────────────────────────
st.title("💡 인사이트 분석")
st.caption("주간 펌웨어 테스트 결과에서 자동 도출된 인사이트")
st.markdown("---")

# 데이터 로드
df = get_current_df()
fail_df = df[df['Result'] == 'FAIL']

# 사이드바
st.sidebar.title("📋 인사이트 분석")
show_help = st.sidebar.checkbox("💡 메뉴 도움말 표시", value=True)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📦 데이터 정보")
st.sidebar.markdown(f"- 전체 테스트: **{len(df):,}**건")
st.sidebar.markdown(f"- Fail 건수: **{len(fail_df):,}**건")
st.sidebar.markdown(f"- 분석 모델: **{df['Model'].nunique()}**개")
st.sidebar.markdown(f"- 분석 FW: **{df['FW_Version'].nunique()}**개")


# 인사이트 분석 메뉴 본문
st.markdown("**필터·차트별 인사이트 + 월별 추이 + 통계 분석을 한 화면에서 확인합니다.**")

# 3개 탭
tab1, tab2, tab3 = st.tabs([
    "🎯 필터 분석",
    "📊 월별·고객사 분석",
    "📈 통계 분석",
])

with tab1:
    if show_help:
        show_menu_help(
            "🎯 필터 분석",
            "IC, Model, Build, Test_Item, Fail_Type 필터로 데이터를 좁혀가며 차트와 인사이트를 확인합니다.",
            [
                "사이드바 필터로 IC/모델/빌드/테스트 항목/결함 종류를 선택하면 차트가 자동 갱신됩니다",
                "차트 바로 아래의 💡 인사이트 박스에서 '한 문장 발견 + 다음 액션'을 자동 생성합니다",
                "인사이트 박스 아래 '📁 관련 영상 목록 보기' 클릭 시 관련 결함 영상을 펼쳐서 볼 수 있습니다",
                "페이지 맨 아래에 카이제곱 통계 검증 결과가 함께 표시됩니다 (전체 데이터일 때만)",
            ]
        )
    st.markdown("---")
    render_filter_analysis(df, fail_df)

with tab2:
    if show_help:
        show_menu_help(
            "📊 월별·고객사 분석",
            "월별 Fail율 추이와 고객사·IC별 결함 분포를 분석합니다.",
            [
                "월별로 Fail율이 어떻게 변화하는지 추이 차트로 확인 가능",
                "고객사·IC별 결함 분포로 어느 그룹이 가장 시급한지 파악",
                "차트마다 자동 인사이트로 의사결정 근거 제공",
            ]
        )
    st.markdown("---")
    render_full_insights(df, fail_df)

with tab3:
    if show_help:
        show_menu_help(
            "📈 통계 분석",
            "카이제곱 검정, Bubble chart, Test_Item × Fail_Type 매트릭스 등 심화 통계 분석",
            [
                "카이제곱 독립성 검정으로 합격률 vs 결함 종류의 진짜 관계 검증",
                "Bubble chart로 Fail_Type별 빈도×심각도 위험도 시각화",
                "Test_Item × Fail_Type 매트릭스로 시나리오별 핫스팟 검출",
            ]
        )
    st.markdown("---")
    render_full_statistics(df, fail_df)
