"""
app_preservation.py - 보존력 시험 OCR (Azure 기반)
상단 버튼 레이아웃: 스킨케어 팀 방식
하단 데이터 표시: 보존력 시험 전용
"""

import streamlit as st
import pandas as pd
import os
import sys
import tempfile
import uuid
from pathlib import Path
from datetime import datetime
import io
import fitz
import copy
import logging
import plotly.graph_objects as go
from PIL import Image

# 프로젝트 루트 추가
current_dir = Path(__file__).parent
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

# ========================================
# 로그 설정 (Streamlit 앱용)
# ========================================
def setup_app_logging():
    """Streamlit 앱 로그 설정"""
    
    # 로그 디렉토리 생성
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # 로그 파일명
    log_filename = os.path.join(
        log_dir,
        f"app_preservation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )
    
    # 로거 설정
    logger = logging.getLogger('app_preservation')
    logger.setLevel(logging.INFO)
    
    # 🔧 중복 출력 방지: 상위 로거로 전파 차단
    logger.propagate = False
    
    # 기존 핸들러 제거
    if logger.hasHandlers():
        logger.handlers.clear()
    
    # 포맷 설정
    formatter = logging.Formatter(
        '%(asctime)s | %(levelname)-8s | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 파일 핸들러
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # 콘솔 핸들러
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    logger.info("="*80)
    logger.info("🌐 Streamlit 앱 시작")
    logger.info(f"📁 앱 로그 파일: {log_filename}")
    logger.info("="*80)
    
    return logger

# 앱 로거 초기화
app_logger = setup_app_logging()

# 🆕 Azure 기반 백엔드 import
from backend import PDFProcessor
from backend_preservation import (
    process_preservation_page,
    PreservationExcelSaver,
    STRAINS
)

# ========================================
# 페이지 설정
# ========================================
st.set_page_config(
    page_title="보존력 시험 OCR 도구",
    layout="wide",
    initial_sidebar_state="collapsed"
)

MAX_PDF_PAGES = 50
MAX_FILE_SIZE_MB = 40

# ========================================
# 세션 상태 초기화
# ========================================
if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())

if "ocr_data_frames" not in st.session_state:
    st.session_state.ocr_data_frames = {}

if "current_page" not in st.session_state:
    st.session_state.current_page = 1

if "saved_pages" not in st.session_state:
    st.session_state.saved_pages = set()

if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = None

if "current_file_bytes" not in st.session_state:
    st.session_state.current_file_bytes = None

if "confirm_reset" not in st.session_state:
    st.session_state.confirm_reset = False

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = {}

# 🆕 Excel Saver 초기화
if "excel_saver" not in st.session_state:
    temp_dir = tempfile.gettempdir()
    excel_path = os.path.join(temp_dir, f"보존력시험_{st.session_state.session_id}.xlsx")
    st.session_state.excel_saver = PreservationExcelSaver(excel_path)
    st.session_state.excel_path = excel_path

# ========================================
# 저장 함수
# ========================================
def save_current_page():
    """현재 페이지 데이터 Excel 저장"""
    key = (st.session_state.current_file_name, st.session_state.current_page)
    
    if key not in st.session_state.ocr_data_frames:
        return True
    
    bundle = st.session_state.ocr_data_frames[key]
    data = bundle.get('data', [])
    date_info = bundle.get('date_info', {})
    
    if not data:
        return True
    
    # 🆕 임시 저장소에서 edited_df 가져오기
    temp_df = st.session_state.get(f'_temp_edited_df_{key}')
    
    if temp_df is not None and len(temp_df) > 0:
        # DataFrame을 딕셔너리 리스트로 변환
        edited_data = temp_df.to_dict('records')
        bundle['data'] = edited_data
    
    # 🆕 편집된 날짜 정보 가져오기
    temp_date = st.session_state.get(f'_temp_edited_date_{key}')
    
    if temp_date is not None:
        date_info = temp_date.copy()
        bundle['date_info'] = date_info
    
    # Excel 저장
    with st.spinner('저장 중...'):
        success = st.session_state.excel_saver.add_test_data(
            test_data=bundle['data'],
            date_info=date_info
        )
    
    if success:
        st.session_state.saved_pages.add(key)
        return True
    else:
        st.error('저장 실패. 다시 시도해주세요.')
        return False

# ========================================
# CSS 스타일
# ========================================
st.markdown("""
<style>
    .compact-header {
        background: linear-gradient(90deg, #0066cc 0%, #0099ff 100%) !important;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        color: white !important;
        margin-bottom: 1rem;
    }
    .status-bar {
        background-color: #f0f2f6 !important;
        padding: 0.5rem;
        border-radius: 5px;
        margin: 0.5rem 0;
        color: #000000 !important;
    }
    
    [data-testid="stAppViewContainer"] .compact-header {
        background: linear-gradient(90deg, #0066cc 0%, #0099ff 100%) !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# ========================================
# 헤더
# ========================================
st.markdown("""
<div class="compact-header" style="background: linear-gradient(90deg, #0066cc 0%, #0099ff 100%) !important; color: white !important;">
    <h1 style="color: white !important; margin: 0 !important;">보존력 시험 OCR 도구</h1>
    <p style="color: white !important; margin: 0 !important;">Azure Document Intelligence 기반 PDF to Excel 자동 변환</p>
</div>
""", unsafe_allow_html=True)

# ========================================
# 파일 업로드 영역
# ========================================
header_col1, header_col2 = st.columns([4, 1])

with header_col1:
    has_work = len(st.session_state.ocr_data_frames) > 0
    
    if not has_work:
        uploaded_file = st.file_uploader(
            "PDF 파일 선택",
            type=["pdf"],
            accept_multiple_files=False,
            label_visibility="collapsed",
            key="file_uploader"
        )
        
        if uploaded_file:
            file_id = f"{uploaded_file.name}_{len(uploaded_file.getvalue())}"
            
            if st.session_state.current_file_name != uploaded_file.name:
                if file_id not in st.session_state.processed_files:
                    app_logger.info(f"📁 새 파일 업로드: {uploaded_file.name}")
                    
                    with st.spinner("🔐 파일 확인 중..."):
                        original_bytes = uploaded_file.getvalue()
                        
                        # 파일 크기 체크
                        file_size_mb = len(original_bytes) / (1024 * 1024)
                        app_logger.info(f"📊 파일 크기: {file_size_mb:.2f}MB")
                        
                        if file_size_mb > MAX_FILE_SIZE_MB:
                            app_logger.error(f"❌ 파일 크기 초과: {file_size_mb:.1f}MB")
                            st.error(f"파일 크기가 제한을 초과했습니다. ({file_size_mb:.1f}MB / {MAX_FILE_SIZE_MB}MB)")
                            st.stop()
                        
                        # DRM 처리
                        drm_success, processed_bytes, drm_message = PDFProcessor.process_drm_if_needed(original_bytes)
                        if not drm_success:
                            app_logger.error(f"❌ DRM 처리 실패: {drm_message}")
                            st.error(f"파일 처리 실패: {drm_message}")
                            st.stop()
                        
                        # 페이지 수 체크
                        try:
                            doc = fitz.open(stream=processed_bytes, filetype="pdf")
                            page_count = doc.page_count
                            doc.close()
                            
                            app_logger.info(f"📄 페이지 수: {page_count}")
                            
                            if page_count > MAX_PDF_PAGES:
                                app_logger.error(f"❌ 페이지 수 초과: {page_count}")
                                st.error(f"PDF 페이지 수가 제한을 초과했습니다. (최대 {MAX_PDF_PAGES}페이지)")
                                st.info(f"현재 PDF: {page_count}페이지")
                                st.stop()
                            
                        except Exception as e:
                            app_logger.error(f"❌ PDF 열기 실패: {e}")
                            st.error(f"❌ PDF 열기 실패: {e}")
                            st.stop()
                        
                        st.session_state.processed_files[file_id] = {
                            'bytes': processed_bytes,
                            'message': drm_message,
                            'name': uploaded_file.name,
                            'page_count': page_count
                        }
                        
                        if "DRM 처리 완료" in drm_message:
                            app_logger.info(f"✅ DRM 처리 완료 | 총 {page_count} 페이지")
                            st.success(f"{drm_message} | 총 {page_count} 페이지")
                        else:
                            app_logger.info(f"✅ 파일 로드 완료 | 총 {page_count} 페이지")
                            st.success(f"파일 로드 완료 | 총 {page_count} 페이지")
                
                processed_file_info = st.session_state.processed_files[file_id]
                st.session_state.current_file_name = uploaded_file.name
                st.session_state.current_file_bytes = processed_file_info['bytes']
                st.session_state.current_file_id = file_id
                st.session_state.current_page = 1
                st.rerun()

# ========================================
# 새로 시작하기 버튼
# ========================================
with header_col2:
    if has_work:
        if not st.session_state.get('reset_confirm', False):
            if st.button("🔄 새로 시작하기", use_container_width=True, type="secondary"):
                st.session_state.reset_confirm = True
                st.rerun()
        else:
            col1, col2 = st.columns(2)
            with col1:
                if st.button("취소", use_container_width=True, type="secondary"):
                    st.session_state.reset_confirm = False
                    st.rerun()
            with col2:
                if st.button("모두 삭제", use_container_width=True, type="primary"):
                    # Excel 파일 삭제
                    if os.path.exists(st.session_state.excel_path):
                        os.remove(st.session_state.excel_path)
                    
                    # 초기화
                    st.session_state.ocr_data_frames = {}
                    st.session_state.saved_pages = set()
                    st.session_state.current_page = 1
                    st.session_state.current_file_name = None
                    st.session_state.current_file_bytes = None
                    st.session_state.current_file_id = None
                    st.session_state.processed_files = {}
                    st.session_state.reset_confirm = False
                    
                    # 새 Excel 생성
                    new_session_id = str(uuid.uuid4())
                    excel_path = os.path.join(tempfile.gettempdir(), f"보존력시험_{new_session_id}.xlsx")
                    st.session_state.excel_saver = PreservationExcelSaver(excel_path)
                    st.session_state.excel_path = excel_path
                    st.session_state.session_id = new_session_id
                    
                    st.success("초기화 완료")
                    st.rerun()
        
        if st.session_state.get('reset_confirm', False):
            st.warning("모든 작업(PDF, OCR 결과, Excel)이 영구 삭제됩니다!")

# ========================================
# 현재 파일 설정
# ========================================
current_file = None
page_count = 0

if st.session_state.get('current_file_bytes'):
    current_file = type('obj', (object,), {
        'name': st.session_state.current_file_name,
        'getvalue': lambda self: st.session_state.current_file_bytes
    })()
    
    page_count = PDFProcessor.extract_page_count(st.session_state.current_file_bytes)
    
    if st.session_state.current_page > page_count:
        st.session_state.current_page = page_count
    if st.session_state.current_page < 1:
        st.session_state.current_page = 1

# ========================================
# 메인 컨텐츠
# ========================================
if current_file:
    st.info("OCR 시작 → 데이터 수정 → 저장 → 다음 페이지 이동 순서로 진행하세요")
    
    # ========================================
    # 상단 액션바 (6개 버튼)
    # ========================================
    action_col1, action_col2, action_col3, action_col4, action_col5, action_col6 = st.columns([2, 2, 2, 2, 1, 2])
    
    # 버튼 1: OCR 시작
    with action_col1:
        key = (current_file.name, st.session_state.current_page)
        ocr_completed = key in st.session_state.ocr_data_frames
        has_data = len(st.session_state.ocr_data_frames.get(key, {}).get('data', [])) > 0
        
        if ocr_completed and has_data:
            button_label = "OCR 완료"
            disabled = True
        elif ocr_completed and not has_data:
            button_label = "OCR 재시도"
            disabled = False
        else:
            button_label = "OCR 시작"
            disabled = False
        
        if st.button(button_label, type="primary", use_container_width=True, disabled=disabled):
            app_logger.info(f"🔍 OCR 시작: {current_file.name} - 페이지 {st.session_state.current_page}")
            
            with st.spinner(f"페이지 {st.session_state.current_page} 처리 중..."):
                result = process_preservation_page(
                    current_file.getvalue(), 
                    st.session_state.current_page - 1
                )
                
                if result['success']:
                    st.session_state.ocr_data_frames[key] = {
                        "data": result['data'],
                        "date_info": result['date_info']
                    }
                    app_logger.info(f"✅ OCR 성공: {len(result['data'])}개 균주 추출")
                    st.success(f"{len(result['data'])}개 균주 데이터 추출 완료")
                    st.rerun()
                else:
                    st.session_state.ocr_data_frames[key] = {
                        "data": [],
                        "date_info": {},
                        "_error": result['message']
                    }
                    app_logger.error(f"❌ OCR 실패: {result['message']}")
                    st.error(f"OCR 실패: {result['message']}")
                    st.info("'OCR 재시도' 버튼을 클릭하여 다시 시도하세요")
                    st.rerun()
    
    # 버튼 2: 이전
    with action_col2:
        if st.button("이전", use_container_width=True, 
                    disabled=(st.session_state.current_page <= 1)):
            st.session_state.current_page -= 1
            st.rerun()
    
    # 버튼 3: 저장
    with action_col3:
        key = (current_file.name, st.session_state.current_page)
        ocr_completed = key in st.session_state.ocr_data_frames
        has_data = len(st.session_state.ocr_data_frames.get(key, {}).get('data', [])) > 0
        
        is_last_page = (st.session_state.current_page >= page_count)
        
        if is_last_page:
            disabled = False
        else:
            disabled = not (ocr_completed and has_data)
        
        if st.button("저장", type="primary", use_container_width=True, disabled=disabled):
            app_logger.info(f"💾 저장 시도: {current_file.name} - 페이지 {st.session_state.current_page}")
            
            if save_current_page():
                if is_last_page:
                    app_logger.info("✅ 마지막 페이지 저장 완료!")
                    st.success("마지막 페이지 저장 완료!")
                else:
                    app_logger.info("✅ 저장 완료!")
                    st.success("저장 완료!")
    
    # 버튼 4: 다음
    with action_col4:
        key = (current_file.name, st.session_state.current_page)
        is_last_page = (st.session_state.current_page >= page_count)
        is_saved = key in st.session_state.saved_pages
        
        disabled = not is_saved or is_last_page
        
        if st.button("다음", type="primary", use_container_width=True, disabled=disabled):
            if not is_last_page:
                st.session_state.current_page += 1
                st.rerun()
    
    # 버튼 5: N/M
    with action_col5:
        saved_count = len(st.session_state.saved_pages)
        st.button(f"{saved_count}/{page_count}", 
                  use_container_width=True, disabled=True)
    
    # 버튼 6: Excel 다운로드
    with action_col6:
        if len(st.session_state.saved_pages) > 0 and os.path.exists(st.session_state.excel_path):
            excel_bytes = st.session_state.excel_saver.get_excel_bytes()
            
            if excel_bytes:
                stats = st.session_state.excel_saver.get_statistics()
                file_size_mb = stats.get('file_size_mb', 0)
                
                st.download_button(
                    label=f"Excel 다운로드 ({file_size_mb:.1f}MB)",
                    data=excel_bytes,
                    file_name=f"보존력시험_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.button("Excel 다운로드", use_container_width=True, disabled=True)
    
    # ========================================
    # 상태 표시줄
    # ========================================
    key = (current_file.name, st.session_state.current_page)
    processed_pages = len(st.session_state.ocr_data_frames)
    
    status_html = f"""
    <div class="status-bar">
        <strong>페이지:</strong> {st.session_state.current_page}/{page_count} | 
        <strong>처리 완료:</strong> {processed_pages}/{page_count}
    </div>
    """
    st.markdown(status_html, unsafe_allow_html=True)
    
    # ========================================
    # 메인 컨텐츠 영역 (2단 레이아웃)
    # ========================================
    left_col, right_col = st.columns([4, 6])

    # 좌측: PDF 미리보기
    with left_col:
        st.markdown("### PDF 미리보기 (마우스 휠/드래그로 조작)")
        
        # PDF 렌더링 (고해상도)
        img_bytes = PDFProcessor.render_page_image(
            current_file.getvalue(), 
            st.session_state.current_page - 1, 
            zoom=3.5  # 고해상도
        )
        
        if img_bytes:
            # Plotly를 이용한 인터랙티브 이미지
            # 이미지 로드
            pil_img = Image.open(io.BytesIO(img_bytes))
            
            # Plotly Figure 생성
            fig = go.Figure()
            
            # 이미지 추가
            fig.add_layout_image(
                dict(
                    source=pil_img,
                    xref="x",
                    yref="y",
                    x=0,
                    y=pil_img.height,
                    sizex=pil_img.width,
                    sizey=pil_img.height,
                    sizing="stretch",
                    layer="below"
                )
            )
            
            # 축 설정
            fig.update_xaxes(
                showgrid=False,
                range=[0, pil_img.width],
                showticklabels=False
            )
            
            fig.update_yaxes(
                showgrid=False,
                range=[0, pil_img.height],
                showticklabels=False,
                scaleanchor="x",
                scaleratio=1
            )
            
            # 레이아웃 설정
            fig.update_layout(
                title=f"페이지 {st.session_state.current_page}/{page_count}",
                width=None,
                height=800,
                margin=dict(l=0, r=0, t=40, b=0),
                xaxis=dict(visible=False),
                yaxis=dict(visible=False),
                hovermode=False,
                dragmode="pan"  # 드래그로 이동
            )
            
            # Plotly 차트 표시
            st.plotly_chart(
                fig,
                use_container_width=True,
                config={
                    'scrollZoom': True,  # 휠 줌
                    'displayModeBar': True,
                    'modeBarButtonsToAdd': ['pan2d', 'zoom2d', 'zoomIn2d', 'zoomOut2d', 'resetScale2d'],
                    'modeBarButtonsToRemove': ['select2d', 'lasso2d']
                }
            )
            
            st.info("💡 **사용법:** 마우스 휠로 확대/축소, 드래그로 이동, 🏠 버튼으로 리셋")
        else:
            st.error("이미지 렌더링 실패")

    # ========================================
    # 우측: OCR 결과 (보존력 시험 전용)
    # ========================================
    with right_col:
        st.markdown("### OCR 결과")
        
        key = (current_file.name, st.session_state.current_page)
        
        # 🆕 자동 OCR (2페이지 이상)
        if key not in st.session_state.ocr_data_frames and st.session_state.current_page > 1:
            with st.spinner("페이지 분석 중... (약 5초 소요)"):
                result = process_preservation_page(
                    current_file.getvalue(), 
                    st.session_state.current_page - 1
                )
                
                if result['success']:
                    st.session_state.ocr_data_frames[key] = {
                        "data": result['data'],
                        "date_info": result['date_info']
                    }
                    st.success(f"자동 OCR 완료: {len(result['data'])}개 균주")
                    st.rerun()
                else:
                    st.session_state.ocr_data_frames[key] = {
                        "data": [],
                        "date_info": {},
                        "_error": result['message']
                    }
                    st.error(f"자동 OCR 실패: {result['message']}")
                    st.info("상단 'OCR 재시도' 버튼으로 다시 시도하세요")
                    st.rerun()
        
        # OCR 결과 표시
        if key in st.session_state.ocr_data_frames:
            bundle = st.session_state.ocr_data_frames[key]
            
            # 에러가 있으면 표시
            if '_error' in bundle:
                st.warning(f"⚠️ 이전 OCR 시도 실패: {bundle['_error']}")
                st.info("데이터를 수정하거나 'OCR 재시도' 버튼을 클릭하세요")
            
            # 데이터가 있으면 표시
            if bundle.get('data'):
                # ========================================
                # 날짜 정보 표시 및 편집
                # ========================================
                
                # 🔧 편집된 날짜가 있으면 우선 사용
                temp_date = st.session_state.get(f'_temp_edited_date_{key}')
                
                if temp_date is not None:
                    date_info = temp_date.copy()
                else:
                    date_info = bundle.get('date_info', {})
                
                # 날짜 정보가 없으면 빈 딕셔너리 생성
                if not date_info or not any(date_info.values()):
                    st.warning("⚠️ 날짜 정보가 없습니다. 직접 입력하세요.")
                    date_info = {
                        'date_0': '',
                        'date_7': '',
                        'date_14': '',
                        'date_28': ''
                    }
                
                st.markdown("**📅 날짜 정보 (편집 가능)**")
                date_df = pd.DataFrame([{
                    '0일': date_info.get('date_0', ''),
                    '7일': date_info.get('date_7', ''),
                    '14일': date_info.get('date_14', ''),
                    '28일': date_info.get('date_28', '')
                }])
                
                # 날짜 에디터 (항상 표시)
                edited_date_df = st.data_editor(
                    date_df,
                    use_container_width=True,
                    height=80,
                    hide_index=True,
                    key=f"date_editor_{current_file.name}_{st.session_state.current_page}",
                    column_config={
                        '0일': st.column_config.TextColumn("0일", help="날짜 형식: MM/DD"),
                        '7일': st.column_config.TextColumn("7일", help="날짜 형식: MM/DD"),
                        '14일': st.column_config.TextColumn("14일", help="날짜 형식: MM/DD"),
                        '28일': st.column_config.TextColumn("28일", help="날짜 형식: MM/DD")
                    }
                )
                
                # 편집된 날짜를 딕셔너리로 변환하여 저장
                if len(edited_date_df) > 0:
                    edited_date_dict = {
                        'date_0': str(edited_date_df.iloc[0]['0일']).strip(),
                        'date_7': str(edited_date_df.iloc[0]['7일']).strip(),
                        'date_14': str(edited_date_df.iloc[0]['14일']).strip(),
                        'date_28': str(edited_date_df.iloc[0]['28일']).strip()
                    }
                    st.session_state[f'_temp_edited_date_{key}'] = edited_date_dict
                
                st.markdown("---")
                
                # ========================================
                # 균주 데이터 테이블
                # ========================================
                st.markdown("**균주 데이터**")
                
                # 🔧 편집된 데이터가 있으면 우선 사용!
                temp_df = st.session_state.get(f'_temp_edited_df_{key}')
                
                if temp_df is not None and len(temp_df) > 0:
                    # 편집된 데이터 사용
                    df = temp_df.copy()
                else:
                    # 원본 데이터 사용
                    data = bundle.get('data', [])
                    if data:
                        df = pd.DataFrame(data)
                    else:
                        df = None
                
                if df is not None and len(df) > 0:
                    
                    # ========================================
                    # 표시용 DataFrame 생성 (검증 이모지 추가)
                    # ========================================
                    df_display = df.copy()
                    
                    # A.brasiliensis 확인 요청 표시
                    def mark_brasiliensis(value, strain):
                        """A.brasiliensis CFU 값에 ⚠️ 추가"""
                        value_str = str(value).strip()
                        
                        if not value_str or value_str == '' or pd.isna(value):
                            return "❌"
                        
                        if 'brasiliensis' in strain.lower():
                            return f"⚠️ {value_str}"
                        
                        return value_str
                    
                    # CFU 컬럼 검증 적용 (❌ 표시)
                    for idx, row in df_display.iterrows():
                        strain = row.get('strain', '')
                        
                        # CFU 컬럼 검증
                        for col in ['cfu_0day', 'cfu_7day', 'cfu_14day', 'cfu_28day']:
                            if col in df_display.columns:
                                df_display.at[idx, col] = mark_brasiliensis(row[col], strain)
                        
                        # 판정 컬럼 검증 (❌ 표시)
                        if 'judgment' in df_display.columns:
                            judgment_val = str(row.get('judgment', '')).strip()
                            if not judgment_val or judgment_val == '' or pd.isna(row.get('judgment')):
                                df_display.at[idx, 'judgment'] = '❌'
                    
                    # ========================================
                    # 중복 제거 (표시용 - 항상 실행!)
                    # ========================================
                    prev_test = None
                    prev_presc = None
                    prev_final = None
                    
                    for i in range(len(df_display)):
                        curr_test = df_display.iloc[i]['test_number']
                        curr_presc = df_display.iloc[i].get('prescription_number', '')
                        curr_final = df_display.iloc[i].get('final_judgment', '')
                        
                        # 시험번호 중복 제거 (❌ 체크 안 함!)
                        if curr_test == prev_test:
                            df_display.at[df_display.index[i], 'test_number'] = ''
                        else:
                            prev_test = curr_test
                        
                        # 처방번호 중복 제거 (❌ 체크 안 함!)
                        if 'prescription_number' in df_display.columns:
                            if curr_presc == prev_presc:
                                df_display.at[df_display.index[i], 'prescription_number'] = ''
                            else:
                                prev_presc = curr_presc
                        
                        # 최종판정 중복 제거 (첫 번째만 표시)
                        if 'final_judgment' in df_display.columns:
                            if curr_final == prev_final and prev_final:
                                df_display.at[df_display.index[i], 'final_judgment'] = ''
                            else:
                                prev_final = curr_final
                    
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
                        'final_judgment': st.column_config.TextColumn("최종판정", width="small", help="시험번호당 첫 번째만")
                    }
                    
                    edited_df = st.data_editor(
                        df_display,
                        column_config=col_config,
                        num_rows="dynamic",
                        hide_index=True,
                        key=f"editor_{current_file.name}_{st.session_state.current_page}",
                        use_container_width=True,
                        height=700
                    )
                    
                    # ========================================
                    # 편집 데이터 정제 (❌, ⚠️ 제거)
                    # ========================================
                    edited_restored = edited_df.copy()
                    
                    # 이모지 제거
                    def remove_emoji(value):
                        value_str = str(value).strip()
                        if value_str == '❌':
                            return ''
                        if '⚠️' in value_str:
                            return value_str.replace('⚠️', '').strip()
                        return value_str
                    
                    for col in ['test_number', 'prescription_number', 'cfu_0day', 'cfu_7day', 'cfu_14day', 'cfu_28day']:
                        if col in edited_restored.columns:
                            edited_restored[col] = edited_restored[col].apply(remove_emoji)
                    
                    # 🔧 빈 값 복원은 원본 데이터일 때만 (temp_df가 없을 때만)
                    if temp_df is None:
                        # 빈 값 복원 (중복 제거된 빈 값을 이전 값으로 채움)
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
                        
                        if 'final_judgment' in edited_restored.columns:
                            prev_final = None
                            for i in range(len(edited_restored)):
                                curr = edited_restored.iloc[i]['final_judgment']
                                if curr == '' or pd.isna(curr):
                                    edited_restored.at[edited_restored.index[i], 'final_judgment'] = prev_final
                                else:
                                    prev_final = curr
                    
                    # 임시 저장소에 저장
                    st.session_state[f'_temp_edited_df_{key}'] = edited_restored
                    
                else:
                    st.info("균주 데이터가 없습니다.")
            else:
                st.info("📋 OCR 데이터가 없습니다")
        
        else:
            st.info("🔍 OCR 시작 버튼을 클릭하여 데이터를 추출하세요")

else:
    st.info("PDF 파일을 업로드하여 시작하세요")