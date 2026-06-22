================================================================
   💡 인사이트 분석 페이지 (단독 실행 버전) — 실행 가이드
================================================================


■ 파일 구성
----------------------------------------------------------------
  app_insight_only.py             ← 메인 파일 (단독 실행)
  SQA_test_results_clean.csv      ← 분석 대상 데이터
  requirements.txt                ← 필요 라이브러리 (선택)


■ 무엇이 들어있나
----------------------------------------------------------------
  본 파일은 SQA 펌웨어 분석 대시보드의 "💡 인사이트 분석" 메뉴만
  독립적으로 실행 가능하도록 추출한 파일입니다.

  포함된 기능:
  - 🎯 필터 분석 탭
      · 5개 필터 (IC/Model/Build/Test_Item/Fail_Type)
      · 차트별 자동 인사이트 (한 문장 + 다음 액션)
      · 인사이트 박스 + 영상 목록 expander
      · 카이제곱 통계 검증 (전체 데이터일 때)
  
  - 📊 월별·고객사 분석 탭
      · 월별 Fail율 추이
      · 고객사·IC별 결함 분포
  
  - 📈 통계 분석 탭
      · 카이제곱 독립성 검정
      · Bubble chart (Fail_Type 위험도)
      · Test_Item × Fail_Type 매트릭스


■ 실행 방법
----------------------------------------------------------------

  [사전 준비]
  - Python 3.9 이상 설치
  - 같은 폴더에 SQA_test_results_clean.csv 존재 필수

  [라이브러리 설치]

      pip install streamlit pandas numpy matplotlib seaborn scipy scikit-learn python-pptx openpyxl

  [실행]

      streamlit run app_insight_only.py

  [브라우저 자동 열림]
      http://localhost:8501


■ 필요한 데이터 파일
----------------------------------------------------------------
  파일명: SQA_test_results_clean.csv
  위치: app_insight_only.py와 같은 폴더
  내용: 7,743건의 펌웨어 테스트 결과 데이터

  ※ 데이터 파일이 없으면 사이트가 정상 실행되지 않습니다


■ 의존성 — 원본 app.py에서 추출한 함수들
----------------------------------------------------------------
  공용 헬퍼:
    - render_insight_box           (인사이트 박스 그리기)
    - render_video_filter_panel    (영상 필터 패널)
    - get_video_link               (Google Drive 링크 생성)
    - ensure_video_id              (Video_Link → video_id 변환)
    - setup_korean_font            (한글 폰트 설정)
    - show_menu_help               (도움말 표시)
  
  데이터 매핑:
    - KEYWORD_DOMAIN_MAP           (결함 → 검토 영역 매칭)
    - get_action_for_keyword       (Action Item 자동 매칭)
  
  인사이트 함수:
    - insight_build_trend          (빌드별 추세)
    - insight_ic_failrate          (IC별 결함 분석)
    - insight_failtype_impact      (Fail_Type 영향 분석)
    - insight_chisquare            (카이제곱 검정)
    - insight_monthly_trend        (월별 추이)
    - insight_customer_analysis    (고객사 분석)
    - insight_test_item_matrix     (Test_Item 매트릭스)
  
  메인 렌더링:
    - render_filter_analysis       (Tab 1)
    - render_full_insights         (Tab 2)
    - render_full_statistics       (Tab 3)


■ 원본과의 차이점
----------------------------------------------------------------
  포함 ✅:
    - 인사이트 분석 메뉴 전체 (Tab 1, 2, 3)
    - 모든 인사이트 함수
    - 영상 필터 패널
    - 카이제곱 통계 검증
  
  제외 ❌:
    - 🏠 홈 화면
    - 🔮 알람 & 예측 메뉴
    - 📥 데이터 & 보고서 메뉴
    - 데이터 업로드 기능 (CSV 파일만 로드)
    - PPT 보고서 자동 생성 기능
    - 회귀 알람 함수


■ 메인 사이트 (전체 기능 포함)
----------------------------------------------------------------
  배포 URL: https://joyjoo.streamlit.app
  GitHub:   https://github.com/heejoocho/sqa-analysis-dashboard


================================================================
                  A74072 조희주 · 서강대학교 AI·SW대학원
================================================================
