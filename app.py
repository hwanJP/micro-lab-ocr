"""
app.py 수정 버전
"""

import streamlit as st
import pandas as pd
import uuid
from datetime import datetime
import os
import tempfile

# 백엔드 모듈 import
from backend import (
    PDFProcessor,
    process_pdf_page,
    ExcelIncrementalSaver,  # 🆕 추가
    STRAINS,
    logger
)

# 페이지 설정
st.set_page_config(
    page_title="보존력 시험 OCR 도구",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 세션 상태 초기화
if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())

if "ocr_data_frames" not in st.session_state:
    st.session_state.ocr_data_frames = {}

if "current_page" not in st.session_state:
    st.session_state.current_page = 1

# 🆕 Excel 증분 저장 객체 초기화
if "excel_saver" not in st.session_state:
    temp_dir = tempfile.gettempdir()
    excel_path = os.path.join(temp_dir, f"보존력시험_{st.session_state.session_id}.xlsx")
    st.session_state.excel_saver = ExcelIncrementalSaver(
        output_path=excel_path,
        template_file=None
    )
    st.session_state.excel_path = excel_path

# CSS 스타일
st.markdown("""
<style>
    .compact-header {
        background: linear-gradient(90deg, #0066cc 0%, #0099ff 100%);
        padding: 0.5rem 1rem;
        border-radius: 5px;
        color: white;
        margin-bottom: 1rem;
    }
    .compact-header h1 {
        font-size: 1.5rem;
        margin: 0;
        padding: 0;
    }
    .compact-header p {
        font-size: 0.9rem;
        margin: 0;
        padding: 0;
        opacity: 0.9;
    }
    
    /* 좌우 컬럼에 직접 스타일 적용 */
    [data-testid="column"] > div > div {
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 1rem;
        background: white;
        min-height: 700px;
    }
    
    .status-bar {
        background: #f8f9fa;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        margin: 0.5rem 0;
        font-size: 0.9rem;
    }
    
    .warning-box {
        background: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 0.75rem;
        margin: 0.5rem 0;
        border-radius: 4px;
    }
    
    .info-section {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 5px;
        margin: 0.5rem 0;
    }
    
    .step-number {
        display: inline-block;
        background: #0066cc;
        color: white;
        width: 24px;
        height: 24px;
        border-radius: 50%;
        text-align: center;
        line-height: 24px;
        font-weight: bold;
        margin-right: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# 헤더
st.markdown("""
<div class="compact-header">
    <h1>보존력 시험 OCR 도구</h1>
    <p>업스테이지 OCR 기반 PDF to Excel 자동 변환</p>
</div>
""", unsafe_allow_html=True)

# 파일 업로드
uploaded_files = st.file_uploader(
    "PDF 파일 선택",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

# 현재 파일 및 페이지 설정
current_file = None
page_count = 0

if uploaded_files:
    file_names = [f.name for f in uploaded_files]
    if len(file_names) > 1:
        selected_file_name = st.selectbox("현재 파일", file_names, label_visibility="collapsed")
    else:
        selected_file_name = file_names[0]
        st.info(f"선택된 파일: {selected_file_name}")
    
    current_file = next(f for f in uploaded_files if f.name == selected_file_name)
    page_count = PDFProcessor.extract_page_count(current_file.getvalue())
    
    if st.session_state.current_page > page_count:
        st.session_state.current_page = page_count
    if st.session_state.current_page < 1:
        st.session_state.current_page = 1

# 데이터 검증 함수
def validate_data(df):
    """데이터 검증"""
    issues = []
    
    if df.empty:
        return issues
    
    missing_test = df[df['test_number'].isna() | (df['test_number'] == '')]
    if not missing_test.empty:
        issues.append(f"시험번호 누락: {len(missing_test)}건")
    
    missing_prescription = df[df['prescription_number'].isna() | (df['prescription_number'] == '')]
    if not missing_prescription.empty:
        issues.append(f"처방번호 누락: {len(missing_prescription)}건")
    
    return issues

# 메인 컨텐츠
if current_file:
    # 상단 액션바
    action_col1, action_col2, action_col3, action_col4, action_col5 = st.columns([2, 2, 2, 1, 2])
    
    with action_col1:
        if st.button("OCR 시작", type="primary", use_container_width=True):
            with st.spinner(f"페이지 {st.session_state.current_page} 처리 중..."):
                result = process_pdf_page(current_file.getvalue(), st.session_state.current_page - 1)
                
                if result['success']:
                    key = (current_file.name, st.session_state.current_page)
                    df_table = pd.DataFrame(result['data'])
                    df_date = pd.DataFrame([result['date_info']]) if result['date_info'] else pd.DataFrame()
                    
                    st.session_state.ocr_data_frames[key] = {"table": df_table, "date": df_date}
                    
                    st.success(result['message'])
                    st.rerun()
                else:
                    st.error(f"처리 실패: {result['message']}")
    
    with action_col2:
        key = (current_file.name, st.session_state.current_page)
        if key in st.session_state.ocr_data_frames:
            if st.button("OCR결과 수정 완료", use_container_width=True):
                # 🆕 즉시 Excel에 저장
                bundle = st.session_state.ocr_data_frames[key]
                df_table = bundle.get("table", pd.DataFrame())
                df_date = bundle.get("date", pd.DataFrame())
                
                # 날짜 정보 추출
                date_info = {}
                if not df_date.empty:
                    date_row = df_date.iloc[0]
                    date_info = {
                        'date_0': date_row.get('date_0', ''),
                        'date_7': date_row.get('date_7', ''),
                        'date_14': date_row.get('date_14', ''),
                        'date_28': date_row.get('date_28', '')
                    }
                
                # Excel 증분 저장
                success = st.session_state.excel_saver.add_test_data(df_table, date_info)
                
                if success:
                    st.success("수정 사항이 Excel에 저장되었습니다")
                    
                    # 저장된 시트 목록 표시
                    sheet_list = st.session_state.excel_saver.get_sheet_list()
                    if sheet_list:
                        st.info(f"저장된 시트: {len(sheet_list)}개")
                else:
                    st.error("Excel 저장 실패")
                
                st.rerun()
        else:
            st.button("OCR결과 수정 완료", use_container_width=True, disabled=True)
    
    with action_col3:
        # Excel 생성 버튼은 유지 (기존 방식과 호환)
        if st.session_state.ocr_data_frames:
            if st.button("Excel 생성", use_container_width=True):
                with st.spinner("Excel 생성 중..."):
                    all_dfs = []
                    for (file_name, page_num), bundle in st.session_state.ocr_data_frames.items():
                        if isinstance(bundle, pd.DataFrame):
                            df_copy = bundle.copy()
                        else:
                            df_copy = bundle.get("table", pd.DataFrame()).copy()
                        all_dfs.append(df_copy)
                    
                    if all_dfs:
                        combined_df = pd.concat(all_dfs, ignore_index=True)
                        data_list = combined_df.to_dict('records')
                        excel_bytes = ExcelGenerator.create_excel(data_list)
                        
                        if excel_bytes:
                            st.session_state['combined_excel_data'] = excel_bytes
                            st.success("Excel 생성 완료")
                        else:
                            st.error("Excel 생성 실패")
        else:
            st.button("Excel 생성", use_container_width=True, disabled=True)
    
    with action_col4:
        if st.button("다음", use_container_width=True):
            if st.session_state.current_page < page_count:
                st.session_state.current_page += 1
                st.rerun()
    
    with action_col5:
        # 🆕 증분 저장된 Excel 다운로드 (우선)
        if os.path.exists(st.session_state.excel_path):
            excel_bytes = st.session_state.excel_saver.get_excel_bytes()
            if excel_bytes:
                st.download_button(
                    label="Excel 다운로드",
                    data=excel_bytes,
                    file_name=f"보존력시험_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        # 기존 방식 Excel도 지원
        elif 'combined_excel_data' in st.session_state:
            st.download_button(
                label="Excel 다운로드",
                data=st.session_state['combined_excel_data'],
                file_name=f"보존력시험_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.button("Excel 다운로드", use_container_width=True, disabled=True)
    
    # 상태 표시줄
    key = (current_file.name, st.session_state.current_page)
    processed_pages = len(st.session_state.ocr_data_frames)
    
    status_html = f"""
    <div class="status-bar">
        <strong>페이지:</strong> {st.session_state.current_page}/{page_count} | 
        <strong>처리 완료:</strong> {processed_pages}/{page_count}
    </div>
    """
    st.markdown(status_html, unsafe_allow_html=True)
    
    # 데이터 검증 경고
    if key in st.session_state.ocr_data_frames:
        bundle = st.session_state.ocr_data_frames[key]
        if not isinstance(bundle, pd.DataFrame):
            df_check = bundle.get("table", pd.DataFrame())
            issues = validate_data(df_check)
            
            if issues:
                warning_html = f"""
                <div class="warning-box">
                    <strong>주의:</strong> {', '.join(issues)}
                </div>
                """
                st.markdown(warning_html, unsafe_allow_html=True)
    
    # 좌우 레이아웃 (4:6 비율)
    left_col, right_col = st.columns([4, 6], gap="medium")
    
    # 좌측: PDF 미리보기
    with left_col:
        st.markdown("#### PDF 미리보기")
        
        img_bytes = PDFProcessor.render_page_image(
            current_file.getvalue(), 
            st.session_state.current_page - 1, 
            zoom=2.5  # zoom 증가
        )
        
        if img_bytes:
            st.image(
                img_bytes,
                use_container_width=True,
                caption=f"{current_file.name} - 페이지 {st.session_state.current_page}/{page_count}"
            )
        else:
            st.error("이미지 렌더링 실패")
    
    # 우측: OCR 결과
    with right_col:
        st.markdown("#### OCR 결과 데이터")
        
        key = (current_file.name, st.session_state.current_page)
        
        if key in st.session_state.ocr_data_frames:
            bundle = st.session_state.ocr_data_frames[key]
            
            if isinstance(bundle, pd.DataFrame):
                df_table = bundle
                df_date = pd.DataFrame(columns=['date_0', 'date_7', 'date_14', 'date_28'])
            else:
                df_table = bundle.get("table", pd.DataFrame())
                df_date = bundle.get("date", pd.DataFrame())
            
            # 날짜 정보
            if not df_date.empty and any(df_date.iloc[0].notna()):
                st.markdown("**날짜 정보**")
                date_display = df_date.copy()
                date_display.columns = ['0일', '7일', '14일', '28일']
                st.dataframe(date_display, use_container_width=True, height=80)
                st.markdown("---")
            
            # 데이터 테이블
            if not df_table.empty:
                col_config = {
                    'test_number': st.column_config.TextColumn("시험번호", width="small"),
                    'prescription_number': st.column_config.TextColumn("처방번호", width="medium"),
                    'strain': st.column_config.SelectboxColumn("균주", options=STRAINS, width="small"),
                    'cfu_0day': st.column_config.TextColumn("0일 CFU", width="small"),
                    'cfu_7day': st.column_config.TextColumn("7일 CFU", width="small"),
                    'cfu_14day': st.column_config.TextColumn("14일 CFU", width="small"),
                    'cfu_28day': st.column_config.TextColumn("28일 CFU", width="small"),
                    'judgment': st.column_config.SelectboxColumn("판정", options=['적합', '부적합'], width="small"),
                    'final_judgment': st.column_config.SelectboxColumn("최종판정", options=['적합', '부적합'], width="small")
                }
                
                edited_df = st.data_editor(
                    df_table,
                    column_config=col_config,
                    num_rows="dynamic",
                    hide_index=True,
                    key=f"editor_{current_file.name}_{st.session_state.current_page}",
                    use_container_width=True,
                    height=500
                )
                
                # 편집된 데이터 저장
                st.session_state.ocr_data_frames[key] = {"table": edited_df, "date": df_date}
                
                # 통계
                st.markdown("---")
                stat_col1, stat_col2, stat_col3 = st.columns(3)
                with stat_col1:
                    st.metric("총 데이터", len(edited_df))
                with stat_col2:
                    st.metric("시험번호", edited_df['test_number'].nunique())
                with stat_col3:
                    st.metric("균주 종류", edited_df['strain'].nunique())
                
            else:
                st.info("OCR 결과 데이터가 없습니다. OCR 시작 버튼을 클릭하세요.")
        
        else:
            st.info("OCR 결과 데이터가 없습니다. OCR 시작 버튼을 클릭하세요.")
    
    # 🆕 하단에 저장된 시트 목록 표시
    if st.session_state.excel_saver:
        sheet_list = st.session_state.excel_saver.get_sheet_list()
        if sheet_list:
            st.markdown("---")
            st.markdown("### 저장된 시트 목록")
            
            cols = st.columns(3)
            for i, sheet_name in enumerate(sheet_list):
                with cols[i % 3]:
                    st.markdown(f"- {sheet_name}")
    
    # 하단 통계
    st.markdown("---")
    st.markdown("### 전체 현황")
    
    def _bundle_len(b):
        try:
            if isinstance(b, pd.DataFrame):
                return len(b)
            table = b.get("table") if isinstance(b, dict) else None
            return len(table) if isinstance(table, pd.DataFrame) else 0
        except Exception:
            return 0
    
    total_records = sum(_bundle_len(b) for b in st.session_state.ocr_data_frames.values())
    
    file_stats = {}
    for (file_name, page_num), bundle in st.session_state.ocr_data_frames.items():
        if file_name not in file_stats:
            file_stats[file_name] = {'pages': 0, 'records': 0}
        file_stats[file_name]['pages'] += 1
        file_stats[file_name]['records'] += _bundle_len(bundle)
    
    stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)
    
    with stats_col1:
        st.metric("처리된 페이지", processed_pages)
    with stats_col2:
        st.metric("추출된 데이터", total_records)
    with stats_col3:
        st.metric("처리된 파일", len(file_stats))
    with stats_col4:
        avg_per_page = round(total_records / processed_pages, 1) if processed_pages > 0 else 0
        st.metric("페이지당 평균", f"{avg_per_page}개")

else:
    # 초기 화면
    st.info("PDF 파일을 업로드하여 시작하세요")
    
    # 사용 방법 (Expander)
    with st.expander("사용 방법 보기", expanded=False):
        st.markdown("""
        <div class="info-section">
            <h4>작업 순서</h4>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            <span class="step-number">1</span><strong>파일 업로드</strong><br>
            상단 파일 선택 영역에서 PDF 파일을 업로드합니다.
            여러 파일을 동시에 선택할 수 있습니다.
            
            <br><br>
            
            <span class="step-number">2</span><strong>OCR 시작</strong><br>
            'OCR 시작' 버튼을 클릭하여 현재 페이지의 데이터를 자동으로 추출합니다.
            업스테이지 AI가 표 형식의 데이터를 인식합니다.
            
            <br><br>
            
            <span class="step-number">3</span><strong>데이터 검토 및 수정</strong><br>
            우측 OCR 결과 테이블에서 추출된 데이터를 확인합니다.
            잘못 인식된 부분은 직접 클릭하여 수정할 수 있습니다.
            행을 추가하거나 삭제할 수도 있습니다.
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <span class="step-number">4</span><strong>수정 완료</strong><br>
            데이터 수정이 끝나면 'OCR결과 수정 완료' 버튼을 클릭하여 
            현재 페이지의 데이터를 Excel 파일에 즉시 저장합니다.
            
            <br><br>
            
            <span class="step-number">5</span><strong>다음 페이지로 이동</strong><br>
            '다음' 버튼을 클릭하여 다음 페이지로 이동합니다.
            2~4단계를 반복하여 모든 페이지를 처리합니다.
            
            <br><br>
            
            <span class="step-number">6</span><strong>Excel 다운로드</strong><br>
            언제든지 'Excel 다운로드' 버튼을 클릭하여 
            지금까지 저장된 데이터를 Excel 파일로 다운로드할 수 있습니다.
            """, unsafe_allow_html=True)
    
    # 주요 기능 (Expander)
    with st.expander("주요 기능 안내", expanded=False):
        st.markdown("""
        <div class="info-section">
            <h4>시스템 기능</h4>
        </div>
        """, unsafe_allow_html=True)
        
        feature_col1, feature_col2, feature_col3 = st.columns(3)
        
        with feature_col1:
            st.markdown("""
            **자동 데이터 추출**
            
            - 시험번호 자동 인식
            - 처방번호 자동 인식
            - 균주명 자동 정규화
            - CFU 값 자동 추출
            - 판정 자동 추출
            """)
        
        with feature_col2:
            st.markdown("""
            **자동 보정 기능**
            
            - OCR 오인식 자동 수정
            - CFU 값 표기 통일
            - 특수문자 정리
            - 균주별 시점별 보정
            - I/1 OCR 오류 보정
            """)
        
        with feature_col3:
            st.markdown("""
            **데이터 검증**
            
            - 시험번호 누락 감지
            - 처방번호 누락 감지
            - 실시간 경고 메시지
            - CFU 값 Log 변환
            - 증분 저장 (데이터 안전)
            """)
        
        st.markdown("---")
        
        st.markdown("""
        <div class="info-section">
            <h4>지원 데이터 형식</h4>
        </div>
        """, unsafe_allow_html=True)
        
        format_col1, format_col2 = st.columns(2)
        
        with format_col1:
            st.markdown("""
            **시험번호 형식**
            - 25E15I14
            - 26E15I14
            - 25A20I02 (A-L 지원)
            
            **처방번호 형식**
            - GB1919-ZMB
            - CCA21201-VAA
            - CC2132-AZLY1
            """)
        
        with format_col2:
            st.markdown("""
            **지원 균주**
            - E.coli (대장균)
            - P.aeruginosa (녹농균)
            - S.aureus (황색포도상구균)
            - C.albicans (칸디다균)
            - A.brasiliensis (아스퍼질러스)
            """)