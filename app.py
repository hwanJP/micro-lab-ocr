"""
app.py 수정 버전
"""

import streamlit as st
import pandas as pd
import os
import sys
import tempfile
import uuid
import logging
from pathlib import Path
from datetime import datetime
import io

# 🆕 PyMuPDF import 추가
import fitz  # PyMuPDF

# 프로젝트 루트를 Python 경로에 추가
current_dir = Path(__file__).parent
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

# 백엔드 모듈 import
from backend import (
    PDFProcessor,
    process_pdf_page,
    ExcelIncrementalSaver,  # 🆕 추가
    STRAINS,
    FallbackManager,
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

if "saved_pages" not in st.session_state:
    st.session_state.saved_pages = set()
    

# 🆕 마지막 날짜 정보 저장
if "last_date_info" not in st.session_state:
    st.session_state.last_date_info = {}
    
# 🆕 페이지별 fallback 관리자
if "fallback_manager" not in st.session_state:
    from backend import FallbackManager
    st.session_state.fallback_manager = FallbackManager()

# 🆕 파일 관련 세션
if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = None

if "current_file_bytes" not in st.session_state:
    st.session_state.current_file_bytes = None

if "confirm_reset" not in st.session_state:
    st.session_state.confirm_reset = False
    
# 세션 초기화에 추가
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = {}  # {file_name: processed_bytes}

if "excel_saver" not in st.session_state:
    temp_dir = tempfile.gettempdir()
    excel_path = os.path.join(temp_dir, f"보존력시험_{st.session_state.session_id}.xlsx")
    st.session_state.excel_saver = ExcelIncrementalSaver(
        output_path=excel_path,
        template_file=None
    )
    st.session_state.excel_path = excel_path
# CSS 스타일
# CSS 스타일 - 최소화 버전
st.markdown("""
<style>
    /* 헤더만 유지 */
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
    
    /* 상태 표시줄 */
    .status-bar {
        background: #f8f9fa;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        margin: 0.5rem 0;
        font-size: 0.9rem;
    }
    
    /* 경고 박스 */
    .warning-box {
        background: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 0.75rem;
        margin: 0.5rem 0;
        border-radius: 4px;
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

# ==================== 파일 업로드 영역 ====================
header_col1, header_col2 = st.columns([4, 1])

with header_col1:
    # 작업 중인지 확인
    has_work = len(st.session_state.ocr_data_frames) > 0
    
    if not has_work:
        # 작업 전: 파일 업로드 가능
        uploaded_file = st.file_uploader(
            "PDF 파일 선택",
            type=["pdf"],
            accept_multiple_files=False,
            label_visibility="collapsed",
            key="file_uploader"
        )
        
        if uploaded_file:
            # 🆕 파일 식별자 생성 (이름 + 크기)
            file_id = f"{uploaded_file.name}_{len(uploaded_file.getvalue())}"
            
            # 🆕 파일이 변경되었는지 확인
            if st.session_state.current_file_name != uploaded_file.name:
                # 🆕 이미 처리된 파일인지 확인
                if file_id not in st.session_state.processed_files:
                    with st.spinner("🔐 파일 확인 중..."):
                        # 원본 파일 bytes
                        original_bytes = uploaded_file.getvalue()
                        
                        # DRM 처리 (최초 1회만)
                        drm_success, processed_bytes, drm_message = PDFProcessor.process_drm_if_needed(original_bytes)
                        
                        if not drm_success:
                            st.error(f"❌ 파일 처리 실패: {drm_message}")
                            logger.error(f"DRM 처리 실패: {drm_message}")
                            st.stop()
                        
                        # 🆕 페이지 수 확인
                        try:
                            doc = fitz.open(stream=processed_bytes, filetype="pdf")
                            page_count = doc.page_count
                            doc.close()
                        except Exception as e:
                            st.error(f"❌ PDF 열기 실패: {e}")
                            logger.error(f"PDF 열기 실패: {e}")
                            st.stop()
                        
                        # 🆕 처리된 파일 캐싱
                        st.session_state.processed_files[file_id] = {
                            'bytes': processed_bytes,
                            'message': drm_message,
                            'name': uploaded_file.name,
                            'page_count': page_count
                        }
                        logger.info(f"✅ 파일 처리 완료 및 캐싱: {file_id} ({page_count} 페이지)")
                        
                        # 사용자에게 알림
                        if "DRM 처리 완료" in drm_message or "DRM 해제" in drm_message:
                            st.success(f"✅ {drm_message} | 총 {page_count} 페이지")
                        else:
                            st.success(f"✅ 파일 로드 완료 | 총 {page_count} 페이지")
                else:
                    logger.info(f"✅ 캐시된 파일 사용: {file_id}")
                
                # 🆕 캐시에서 처리된 파일 가져오기
                processed_file_info = st.session_state.processed_files[file_id]
                
                # 세션에 저장
                st.session_state.current_file_name = uploaded_file.name
                st.session_state.current_file_bytes = processed_file_info['bytes']  # ← DRM 해제된 bytes
                st.session_state.current_file_id = file_id  # 🆕 파일 ID 저장
                st.session_state.current_page = 1
                
                logger.info(f"📁 파일 설정 완료: {uploaded_file.name}")
                st.rerun()

with header_col2:
    if has_work:
        # 🆕 새로 시작하기 버튼
        if st.button("🔄 새로 시작하기", use_container_width=True, type="secondary"):
            if st.session_state.get('confirm_reset', False):
                # 전체 초기화
                st.session_state.ocr_data_frames = {}
                st.session_state.saved_pages = set()
                st.session_state.current_page = 1
                st.session_state.last_date_info = {}
                st.session_state.fallback_manager.reset()
                st.session_state.current_file_name = None
                st.session_state.current_file_bytes = None
                st.session_state.current_file_id = None  # 🆕 추가
                st.session_state.confirm_reset = False
                # 🆕 캐시는 유지 (같은 파일 다시 업로드 시 빠르게)
                # st.session_state.processed_files = {}  # 필요시 주석 해제
                
                # Excel 초기화
                temp_dir = tempfile.gettempdir()
                excel_path = os.path.join(temp_dir, f"보존력시험_{st.session_state.session_id}.xlsx")
                st.session_state.excel_saver = ExcelIncrementalSaver(
                    output_path=excel_path,
                    template_file=None
                )
                st.session_state.excel_path = excel_path
                
                logger.info("🔄 전체 초기화 완료")
                st.success("✅ 새로 시작합니다")
                st.rerun()
            else:
                st.session_state.confirm_reset = True
                st.warning("⚠️ 작업 내용이 삭제됩니다. 다시 클릭하면 초기화됩니다.")
                st.rerun()

# 🆕 현재 파일 설정
current_file = None
page_count = 0

if st.session_state.get('current_file_bytes'):
    # 세션에서 파일 로드
    import io
    current_file = type('obj', (object,), {
        'name': st.session_state.current_file_name,
        'getvalue': lambda self: st.session_state.current_file_bytes  # self 추가!
    })()
    
    page_count = PDFProcessor.extract_page_count(st.session_state.current_file_bytes)
    
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
                # 🆕 DRM 처리 상태 표시
                drm_placeholder = st.empty()
                # drm_placeholder.info("🔐 DRM 확인 중...")
                
                result = process_pdf_page(
                    current_file.getvalue(), 
                    st.session_state.current_page - 1,
                    st.session_state.fallback_manager  # 🎯 추가
                )
                
                drm_placeholder.empty()  # DRM 메시지 제거
                
                if result['success']:
                    key = (current_file.name, st.session_state.current_page)
                    df_table = pd.DataFrame(result['data'])
                    df_date_raw = result['date_info']  # 딕셔너리
                    
                    # 🆕 날짜 정보 처리
                    if df_date_raw and any(df_date_raw.values()):
                        # 새로운 날짜 정보가 있으면 저장
                        st.session_state.last_date_info = df_date_raw.copy()
                        df_date = pd.DataFrame([df_date_raw])
                        logger.info(f"📅 새로운 날짜 정보 저장: {df_date_raw}")
                    elif st.session_state.last_date_info:
                        # 날짜 정보가 없으면 이전 값 재사용
                        df_date = pd.DataFrame([st.session_state.last_date_info])
                        logger.info(f"🔄 이전 날짜 정보 재사용: {st.session_state.last_date_info}")
                    else:
                        # 날짜 정보가 전혀 없는 경우
                        df_date = pd.DataFrame()
                        logger.warning("⚠️ 날짜 정보 없음")
                    
                    st.session_state.ocr_data_frames[key] = {"table": df_table, "date": df_date}
                    
                    st.success(result['message'])
                    st.rerun()
                else:
                    st.error(f"처리 실패: {result['message']}")
    
    with action_col2:
        key = (current_file.name, st.session_state.current_page)
        if key in st.session_state.ocr_data_frames:
            if st.button("Excel에 저장", use_container_width=True):
                bundle = st.session_state.ocr_data_frames[key]
                df_table = bundle.get("table", pd.DataFrame())
                df_date = bundle.get("date", pd.DataFrame())
                
                date_info = {}
                if not df_date.empty:
                    date_row = df_date.iloc[0]
                    date_info = {
                        'date_0': date_row.get('date_0', ''),
                        'date_7': date_row.get('date_7', ''),
                        'date_14': date_row.get('date_14', ''),
                        'date_28': date_row.get('date_28', '')
                    }
                
                success = st.session_state.excel_saver.add_test_data(df_table, date_info)
                
                if success:
                    # 🆕 저장 완료 기록
                    st.session_state.saved_pages.add(key)
                    
                    if not df_table.empty and 'test_number' in df_table.columns:
                        test_count = df_table['test_number'].nunique()
                        st.success(f"{test_count}개 시험이 저장되었습니다")
                    else:
                        st.success("저장되었습니다")
                    
                    sheet_list = st.session_state.excel_saver.get_sheet_list()
                    if sheet_list:
                        st.info(f"총 저장된 시트: {len(sheet_list)}개")
                else:
                    st.error("Excel 저장 실패")
                
                st.rerun()
        else:
            st.button("Excel에 저장", use_container_width=True, disabled=True)
    

    with action_col3:
        # 비활성 버튼으로 통계 표시 (평행 정렬)
        if st.session_state.excel_saver:
            stats = st.session_state.excel_saver.get_statistics()
            sheet_count = stats['test_sheets']
        else:
            sheet_count = 0
        
        st.button(f"저장: {sheet_count}개", use_container_width=True, disabled=True)
    
    with action_col4:
        key = (current_file.name, st.session_state.current_page)
        
        # 🆕 저장 여부 확인
        is_saved = key in st.session_state.saved_pages
        has_data = key in st.session_state.ocr_data_frames
        
        # 데이터는 있지만 저장 안된 경우
        if has_data and not is_saved:
            st.button("다음", use_container_width=True, disabled=True)
            st.caption("저장 후 이동")
        # 저장됨 또는 데이터 없음
        else:
            if st.button("다음", use_container_width=True):
                if st.session_state.current_page < page_count:
                    st.session_state.current_page += 1
                    # 🆕 여기 2줄 추가 (시작)
                    st.session_state.fallback_manager.reset()
                    logger.info(f"▶ 페이지 {st.session_state.current_page}로 이동 - Fallback 초기화")

                    st.rerun()
    
    with action_col5:
        # 증분 저장된 Excel 다운로드
        if os.path.exists(st.session_state.excel_path):
            excel_bytes = st.session_state.excel_saver.get_excel_bytes()
            if excel_bytes:
                # 🆕 파일 크기 표시
                stats = st.session_state.excel_saver.get_statistics()
                file_size_mb = stats.get('file_size_mb', 0)
                
                st.download_button(
                    label=f"Excel 다운로드 ({file_size_mb}MB)",
                    data=excel_bytes,
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
                pass
                # warning_html = f"""
                # <div class="warning-box">
                #     <strong>주의:</strong> {', '.join(issues)}
                # </div>
                # """
                # st.markdown(warning_html, unsafe_allow_html=True)
    
    # 좌우 레이아웃 (4:6 비율)
    left_col, right_col = st.columns([4, 6], gap="medium")

    # 좌측: PDF 미리보기
    with left_col:
        # 🆕 네이티브 컨테이너 사용
        with st.container(border=True):
            st.markdown("#### PDF 미리보기")
            
            img_bytes = PDFProcessor.render_page_image(
                current_file.getvalue(), 
                st.session_state.current_page - 1, 
                zoom=2.5
            )
            
            if img_bytes:
                st.image(
                    img_bytes,
                    caption=f"{current_file.name} - 페이지 {st.session_state.current_page}/{page_count}",
                    use_container_width=True
                )
            else:
                st.error("이미지 렌더링 실패")

    # 우측: OCR 결과
    with right_col:
            # 🆕 네이티브 컨테이너 사용
            with st.container(border=True, height=1100):
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
                    
                    # 🆕 날짜 정보 항상 표시
                    if not df_date.empty and any(df_date.iloc[0].notna()):
                        st.markdown("**날짜 정보**")
                        date_display = df_date.copy()
                        date_display.columns = ['0일', '7일', '14일', '28일']
                        st.dataframe(date_display, use_container_width=True, height=80)
                    elif st.session_state.last_date_info:
                        st.markdown("**날짜 정보** (이전 페이지)")
                        date_display = pd.DataFrame([{
                            '0일': st.session_state.last_date_info.get('date_0', ''),
                            '7일': st.session_state.last_date_info.get('date_7', ''),
                            '14일': st.session_state.last_date_info.get('date_14', ''),
                            '28일': st.session_state.last_date_info.get('date_28', '')
                        }])
                        st.dataframe(date_display, use_container_width=True, height=80)
                        st.caption("이전 페이지의 날짜 정보를 사용합니다")
                    else:
                        st.warning("날짜 정보 없음")
                    
                    # 데이터 테이블
                    if not df_table.empty:
                        # 🆕 표시용 DataFrame 생성
                        df_display = df_table.copy()
                        
                        # ========================================
                        # 검증 함수 1: 일반 누락 표시 (기존)
                        # ========================================
                        def mark_missing(value):
                            """누락 표시"""
                            value_str = str(value).strip()
                            if not value_str or value_str == '' or pd.isna(value):
                                return "❌"
                            return value
                        
                        
                        # ========================================
                        # 검증 함수 2: A.brasiliensis 확인 요청 (신규)
                        # ========================================
                        def mark_brasiliensis(value, strain):
                            """
                            A.brasiliensis 확인 요청 표시
                            
                            Args:
                                value: CFU 값
                                strain: 균주명
                                
                            Returns:
                                str: 
                                    - 누락: '❌'
                                    - A.brasiliensis: '⚠️ {값}'
                                    - 기타: '{값}'
                            """
                            value_str = str(value).strip()
                            
                            # 누락
                            if not value_str or value_str == '' or pd.isna(value):
                                return "❌"
                            
                            # A.brasiliensis면 ⚠️ 추가
                            if 'brasiliensis' in strain.lower():
                                return f"⚠️ {value_str}"
                            
                            return value_str
                        
                        
                        # ========================================
                        # 이모지 제거 함수 (저장용)
                        # ========================================
                        def remove_emoji(value):
                            """검증 이모지 제거 (저장용)"""
                            value_str = str(value).strip()
                            
                            if value_str == '❌':
                                return ''
                            
                            if '⚠️' in value_str:
                                return value_str.replace('⚠️', '').strip()
                            
                            return value_str
                        
                        
                        # ========================================
                        # 🆕 CFU 컬럼 검증 적용 (A.brasiliensis 체크)
                        # ========================================
                        for idx, row in df_display.iterrows():
                            strain = row.get('strain', '')
                            
                            for col in ['cfu_0day', 'cfu_7day', 'cfu_14day', 'cfu_28day']:
                                if col in df_display.columns:
                                    df_display.at[idx, col] = mark_brasiliensis(row[col], strain)
                        
                        
                        # ========================================
                        # 중복 제거 + 시험번호/처방번호 누락 표시 (기존)
                        # ========================================
                        prev_test = None
                        prev_presc = None
                        
                        for i in range(len(df_display)):
                            curr_test = df_display.iloc[i]['test_number']
                            curr_presc = df_display.iloc[i].get('prescription_number', '')
                            
                            # 시험번호
                            if curr_test == prev_test:
                                df_display.at[df_display.index[i], 'test_number'] = ''
                            else:
                                test_str = str(curr_test).strip()
                                if not test_str or test_str == '' or pd.isna(curr_test):
                                    df_display.at[df_display.index[i], 'test_number'] = '❌'
                                prev_test = curr_test
                            
                            # 처방번호
                            if 'prescription_number' in df_display.columns:
                                if curr_presc == prev_presc:
                                    df_display.at[df_display.index[i], 'prescription_number'] = ''
                                else:
                                    presc_str = str(curr_presc).strip()
                                    if not presc_str or presc_str == '' or pd.isna(curr_presc):
                                        df_display.at[df_display.index[i], 'prescription_number'] = '❌'
                                    prev_presc = curr_presc
                        
                        
                        # ========================================
                        # 데이터 에디터
                        # ========================================
                        col_config = {
                            'test_number': st.column_config.TextColumn("시험번호", width="small"),
                            'prescription_number': st.column_config.TextColumn("처방번호", width="small"),
                            'strain': st.column_config.SelectboxColumn("균주", options=STRAINS, width="small"),
                            'cfu_0day': st.column_config.TextColumn("0일 CFU", width="small", help="❌=누락, ⚠️=확인필요"),
                            'cfu_7day': st.column_config.TextColumn("7일 CFU", width="small", help="❌=누락, ⚠️=확인필요"),
                            'cfu_14day': st.column_config.TextColumn("14일 CFU", width="small", help="❌=누락, ⚠️=확인필요"),
                            'cfu_28day': st.column_config.TextColumn("28일 CFU", width="small", help="❌=누락, ⚠️=확인필요"),
                            'judgment': st.column_config.SelectboxColumn("판정", options=['적합', '부적합'], width="small"),
                            'final_judgment': st.column_config.SelectboxColumn("최종판정", options=['적합', '부적합'], width="small")
                        }
                        
                        edited_df = st.data_editor(
                            df_display,
                            column_config=col_config,
                            num_rows="dynamic",
                            hide_index=True,
                            key=f"editor_{current_file.name}_{st.session_state.current_page}",
                            use_container_width=True,
                            height=800
                        )
                        
                        
                        # ========================================
                        # 편집 데이터 정제 (❌, ⚠️ 제거)
                        # ========================================
                        edited_restored = edited_df.copy()
                        
                        # 모든 컬럼에서 이모지 제거
                        for col in ['test_number', 'prescription_number', 'cfu_0day', 'cfu_7day', 'cfu_14day', 'cfu_28day']:
                            if col in edited_restored.columns:
                                edited_restored[col] = edited_restored[col].apply(remove_emoji)
                        
                        # 빈 값 복원
                        prev_test = None
                        for i in range(len(edited_restored)):
                            curr = edited_restored.iloc[i]['test_number']
                            if curr == '' or pd.isna(curr):
                                edited_restored.at[edited_restored.index[i], 'test_number'] = prev_test
                            else:
                                prev_test = curr
                        
                        if 'prescription_number' in edited_restored.columns:
                            prev_presc = None
                            for i in range(len(edited_restored)):
                                curr = edited_restored.iloc[i]['prescription_number']
                                if curr == '' or pd.isna(curr):
                                    edited_restored.at[edited_restored.index[i], 'prescription_number'] = prev_presc
                                else:
                                    prev_presc = curr
                        
                        # 편집된 데이터 저장
                        st.session_state.ocr_data_frames[key] = {"table": edited_restored, "date": df_date}
                        
                    else:
                        st.info("OCR 결과 데이터가 없습니다. OCR 시작 버튼을 클릭하세요.")
                
                else:
                    st.info("OCR 결과 데이터가 없습니다. OCR 시작 버튼을 클릭하세요.")
        
        
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