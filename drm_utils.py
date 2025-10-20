"""
DRM 유틸리티 모듈
PDF 파일의 DRM 판별 및 해제 기능 제공
"""

import os
import io
import requests
from typing import Dict, Union, Tuple, Optional
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


# ========================================
# DRM 판별 함수
# ========================================

def detect_drm(file_input: Union[str, io.BytesIO]) -> Dict[str, any]:
    """
    PDF 파일의 DRM 보호 여부를 다각도로 판별
    
    Args:
        file_input: 파일 경로(str) 또는 BytesIO 객체
        
    Returns:
        dict: {
            "is_drm": bool,
            "method": str,
            "confidence": str,
            "details": dict,
            "recommendation": str
        }
    """
    result = {
        "is_drm": False,
        "method": None,
        "confidence": "low",
        "details": {},
        "recommendation": "정상 파일로 추정"
    }
    
    try:
        # 1. PyPDF2로 암호화 확인
        try:
            import PyPDF2
            
            if isinstance(file_input, str):
                f = open(file_input, 'rb')
            else:
                file_input.seek(0)
                f = file_input
            
            reader = PyPDF2.PdfReader(f)
            
            if reader.is_encrypted:
                result["is_drm"] = True
                result["method"] = "PyPDF2 암호화 감지"
                result["confidence"] = "high"
                result["details"]["encrypted"] = True
                result["recommendation"] = "DRM 해제 필요"
                
                if isinstance(file_input, str):
                    f.close()
                return result
            
            if isinstance(file_input, str):
                f.close()
        
        except ImportError:
            logger.warning("PyPDF2 미설치")
        except Exception as e:
            result["details"]["pypdf2_error"] = str(e)
            logger.debug(f"PyPDF2 확인 실패: {e}")
        
        # 2. pdfplumber로 텍스트 추출 시도
        try:
            import pdfplumber
            
            if isinstance(file_input, str):
                pdf = pdfplumber.open(file_input)
            else:
                file_input.seek(0)
                pdf = pdfplumber.open(file_input)
            
            if len(pdf.pages) > 0:
                text = pdf.pages[0].extract_text()
                
                if not text or len(text.strip()) < 10:
                    result["is_drm"] = True
                    result["method"] = "텍스트 추출 실패"
                    result["confidence"] = "medium"
                    result["details"]["text_extractable"] = False
                    result["recommendation"] = "DRM 가능성 있음"
                else:
                    result["details"]["text_extractable"] = True
                    result["details"]["text_length"] = len(text)
            
            pdf.close()
        
        except ImportError:
            logger.warning("pdfplumber 미설치")
        except Exception as e:
            result["is_drm"] = True
            result["method"] = "파일 열기 실패"
            result["confidence"] = "high"
            result["details"]["error"] = str(e)
            result["recommendation"] = "DRM으로 보호된 파일"
            return result
        
        # 3. 바이너리 헤더 확인
        try:
            if isinstance(file_input, str):
                with open(file_input, 'rb') as f:
                    header = f.read(2048)
            else:
                file_input.seek(0)
                header = file_input.read(2048)
                file_input.seek(0)
            
            if not header.startswith(b'%PDF'):
                result["details"]["is_pdf"] = False
                result["recommendation"] = "PDF 파일이 아님"
                return result
            
            result["details"]["is_pdf"] = True
            
            if b'/Encrypt' in header:
                result["is_drm"] = True
                result["method"] = "암호화 플래그 감지"
                result["confidence"] = "high"
                result["details"]["encrypt_flag"] = True
                result["recommendation"] = "DRM 해제 필요"
        
        except Exception as e:
            result["details"]["header_check_error"] = str(e)
            logger.debug(f"헤더 확인 실패: {e}")
        
        return result
    
    except Exception as e:
        logger.error(f"DRM 판별 중 오류: {e}")
        return {
            "is_drm": None,
            "method": "판별 실패",
            "confidence": "unknown",
            "details": {"error": str(e)},
            "recommendation": "수동 확인 필요"
        }


# ========================================
# DRM 해제 함수
# ========================================

def decrypt_drm_file(
    file_input: Union[str, io.BytesIO],
    api_url: str = "https://cnr.kolmar.co.kr/api/services/app/Crypt/FileThirdPartyDecryption",
    api_key: Optional[str] = None
) -> Tuple[bool, Union[bytes, str]]:
    """
    DRM 걸린 파일 해제
    
    Args:
        file_input: 파일 경로(str) 또는 BytesIO 객체
        api_url: DRM 해제 API URL
        api_key: API 인증 키 (필요시)
        
    Returns:
        Tuple[bool, Union[bytes, str]]: (성공여부, 해제된파일bytes or 오류메시지)
    """
    try:
        # 파일 데이터 준비
        if isinstance(file_input, str):
            file_name = os.path.basename(file_input)
            with open(file_input, 'rb') as f:
                file_bytes = f.read()
        else:
            file_name = "uploaded_file.pdf"
            file_input.seek(0)
            file_bytes = file_input.read()
            file_input.seek(0)
        
        logger.info(f"DRM 해제 요청: {file_name} ({len(file_bytes):,} bytes)")
        
        # 멀티파트 폼 데이터
        files = {
            'formFile': (file_name, file_bytes, 'application/pdf')
        }
        
        # 헤더 설정
        headers = {}
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"
        
        # API 호출
        response = requests.post(
            api_url,
            files=files,
            headers=headers,
            timeout=30
        )
        
        logger.info(f"DRM 해제 응답: {response.status_code}")
        
        if response.status_code == 200:
            logger.info(f"DRM 해제 성공 ({len(response.content):,} bytes)")
            return True, response.content
        else:
            error_msg = f"DRM 해제 실패 (HTTP {response.status_code})"
            logger.error(f"{error_msg}: {response.text[:200]}")
            return False, error_msg
    
    except requests.exceptions.ConnectionError as e:
        error_msg = f"연결 오류: SSLVPN 접속 확인 필요"
        logger.error(f"{error_msg}: {e}")
        return False, error_msg
    
    except requests.exceptions.Timeout:
        error_msg = "타임아웃: 30초 내에 응답이 없습니다"
        logger.error(error_msg)
        return False, error_msg
    
    except Exception as e:
        error_msg = f"예외 발생: {e}"
        logger.error(error_msg, exc_info=True)
        return False, error_msg


# ========================================
# 통합 처리 함수
# ========================================

def process_pdf_with_drm(
    file_input: Union[str, io.BytesIO],
    api_url: str = "https://cnr.kolmar.co.kr/api/services/app/Crypt/FileThirdPartyDecryption",
    api_key: Optional[str] = None
) -> Tuple[bool, Union[io.BytesIO, str]]:
    """
    PDF 파일의 DRM을 자동으로 판별하고 해제
    
    Args:
        file_input: 파일 경로(str) 또는 BytesIO 객체
        api_url: DRM 해제 API URL
        api_key: API 인증 키
        
    Returns:
        Tuple[bool, Union[io.BytesIO, str]]: 
            (성공여부, 처리된파일BytesIO or 오류메시지)
    """
    try:
        # 1. DRM 판별
        logger.info("DRM 판별 시작")
        drm_info = detect_drm(file_input)
        
        logger.info(
            f"DRM 판별 결과: is_drm={drm_info['is_drm']}, "
            f"method={drm_info['method']}, "
            f"confidence={drm_info['confidence']}"
        )
        
        # 2. DRM이 없으면 원본 파일 반환
        if not drm_info["is_drm"]:
            logger.info("DRM 없음 - 원본 파일 사용")
            
            if isinstance(file_input, str):
                with open(file_input, 'rb') as f:
                    return True, io.BytesIO(f.read())
            else:
                file_input.seek(0)
                file_bytes = file_input.read()
                file_input.seek(0)
                return True, io.BytesIO(file_bytes)
        
        # 3. DRM 해제 시도
        logger.info("DRM 감지 - 해제 시도")
        decrypt_success, decrypt_result = decrypt_drm_file(
            file_input,
            api_url=api_url,
            api_key=api_key
        )
        
        if decrypt_success:
            logger.info("DRM 해제 완료")
            return True, io.BytesIO(decrypt_result)
        else:
            logger.error(f"DRM 해제 실패: {decrypt_result}")
            return False, decrypt_result
    
    except Exception as e:
        error_msg = f"PDF DRM 처리 중 오류: {e}"
        logger.error(error_msg, exc_info=True)
        return False, error_msg


# ========================================
# Streamlit 업로드 파일 처리
# ========================================

def process_streamlit_uploaded_file(
    uploaded_file,
    api_url: str = "https://cnr.kolmar.co.kr/api/services/app/Crypt/FileThirdPartyDecryption",
    api_key: Optional[str] = None
) -> Tuple[bool, Union[io.BytesIO, str]]:
    """
    Streamlit UploadedFile 객체의 DRM을 처리
    
    Args:
        uploaded_file: Streamlit UploadedFile 객체
        api_url: DRM 해제 API URL
        api_key: API 인증 키
        
    Returns:
        Tuple[bool, Union[io.BytesIO, str]]: 
            (성공여부, 처리된파일BytesIO or 오류메시지)
    """
    try:
        # UploadedFile을 BytesIO로 변환
        uploaded_file.seek(0)
        file_bytes = uploaded_file.read()
        uploaded_file.seek(0)
        
        file_io = io.BytesIO(file_bytes)
        
        # DRM 처리
        return process_pdf_with_drm(file_io, api_url, api_key)
    
    except Exception as e:
        error_msg = f"업로드 파일 처리 중 오류: {e}"
        logger.error(error_msg, exc_info=True)
        return False, error_msg