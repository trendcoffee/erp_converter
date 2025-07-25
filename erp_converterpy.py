import pandas as pd
import os
import streamlit as st
from io import BytesIO # Streamlit 다운로드를 위해 필요

def convert_file(uploaded_file):
    """
    업로드된 엑셀 파일을 받아 이카운트 형식으로 변환하는 함수.
    Args:
        uploaded_file: Streamlit의 st.file_uploader를 통해 업로드된 파일 객체.
    Returns:
        변환된 DataFrame (성공 시) 또는 None (실패 시).
    """
    try:
        # 업로드된 파일을 Pandas DataFrame으로 읽기
        # BytesIO를 사용하여 파일 객체를 직접 전달
        df = pd.read_excel(uploaded_file)

        # NaN 값을 빈 문자열로 일괄 전처리 (문자열 컬럼의 안정성 확보)
        df.fillna('', inplace=True) 

        # 금액, 주문수량, 대표운송장번호 정리 및 숫자 변환 (변환 실패 시 0으로 채움)
        df["금액"] = df["금액"].astype(str).str.replace(",", "")
        df["금액"] = pd.to_numeric(df["금액"], errors="coerce").fillna(0)
        
        df["주문수량"] = pd.to_numeric(df["주문수량"], errors="coerce").fillna(0)
        
        # 대표운송장번호는 NaN 허용하는 정수형으로 변환
        df["대표운송장번호"] = pd.to_numeric(df["대표운송장번호"], errors="coerce").astype(pd.Int64Dtype())
        
        # 주문수량이 0인 행 제거 (기존 로직 유지)
        df = df[df["주문수량"] != 0]

        # 계산
        # 주문수량이 0인 행은 이미 제거되었으므로 ZeroDivisionError 걱정 없음
        단가 = (df["금액"] / df["주문수량"]).fillna(0) 
        수량 = df["주문수량"].fillna(0)
        총금액 = (단가 * 수량).round().fillna(0)
        부가세 = (총금액 / 11).fillna(0).astype(int)
        공급가액 = (총금액 - 부가세).fillna(0).astype(int)

        # 거래처명 및 수집처 조건 함수
        def get_client_name(ch):
            ch = str(ch).strip().upper()
            if ch == "GMKT":
                return "지마켓글로벌 유한책임회사", "지마켓"
            elif ch == "AUCT":
                return "지마켓글로벌 유한책임회사", "옥션"
            elif ch == "SSG":
                return "(주)에스에스지닷컴", "SSG"
            elif ch == "NFA":
                return "네이버파이낸셜 주식회사", "스마트스토어"
            elif ch == "롯데ON":
                return "롯데쇼핑주식회사", "롯데온"
            elif ch == "쿠팡":
                return "쿠팡 주식회사", "쿠팡"
            elif ch == "11번가":
                return "십일번가 주식회사", "11번가"
            else:
                return ch, ch
        
        # '판매채널' 컬럼 존재 여부 확인 후 처리
        if "판매채널" not in df.columns:
            st.error("오류: 필수 컬럼 '판매채널'이(가) 파일에 없습니다.")
            return None # 함수 종료 및 None 반환
            
        거래처명, 수집처 = zip(*df["판매채널"].map(get_client_name))

        생산전표_N_목록 = [
            "YELLOW_NOZZLE_2EA", "WHITE_NOZZLE_2EA", "RED_NOZZLE_2EA",
            "YELLOW_T_NOZZLE-CAP", "WHITE_T_NOZZLE-CAP"
        ]

        result = pd.DataFrame()
        # '출고일자' 컬럼 존재 여부 확인 후 처리
        if "출고일자" not in df.columns:
            st.error("오류: 필수 컬럼 '출고일자'이(가) 파일에 없습니다.")
            return None
        # 출고일자 변환 시 오류 처리 및 NaN 채우기
        result["일자"] = pd.to_datetime(df["출고일자"], errors='coerce').dt.strftime("%Y%m%d").fillna('')
        
        result["순번"] = ""
        result["거래처코드"] = ""
        result["거래처명"] = 거래처명
        result["담당자"] = ""
        result["출하창고"] = "300"
        result["거래유형"] = ""
        result["통화"] = ""
        result["환율"] = ""
        result["잔액"] = ""
        result["참고"] = ""

        # '품목코드', '품목명', '받는사람명', '대표운송장번호' 컬럼 존재 여부 확인
        required_cols_for_result = ["품목코드", "품목명", "받는사람명", "대표운송장번호"]
        for col in required_cols_for_result:
            if col not in df.columns:
                st.error(f"오류: 필수 컬럼 '{col}'이(가) 파일에 없습니다.")
                return None

        result["품목코드"] = df["품목코드"]
        result["품목명"] = df["품목명"]
        result["규격"] = ""
        result["수량"] = 수량
        result["단가"] = 단가
        result["외화금액"] = ""
        result["공급가액"] = 공급가액
        result["부가세"] = 부가세
        result["수집처"] = list(수집처)
        result["수취인"] = df["받는사람명"]
        # 대표운송장번호는 Int64Dtype으로 변환 후 다시 str로 변환
        result["운송장번호"] = df["대표운송장번호"].astype(str).replace('<NA>', '') # <NA>는 Int64Dtype의 NaN 표현
        result["적요"] = ""
        result["생산전표생성"] = result["품목코드"].apply(lambda x: "N" if str(x).strip() in 생산전표_N_목록 else "Y")

        columns_order = [
            "일자", "순번", "거래처코드", "거래처명", "담당자", "출하창고", "거래유형", "통화", "환율", "잔액", "참고",
            "품목코드", "품목명", "규격", "수량", "단가", "외화금액", "공급가액", "부가세",
            "수집처", "수취인", "운송장번호", "적요", "생산전표생성"
        ]
        result = result[columns_order]
        
        # DataFrame이 비어있지 않은지 확인 후 마지막 행 제거
        if not result.empty and '품목코드' in result.columns and result['품목코드'].iloc[-1] == '':
            result = result[:-1]
        
        return result

    except KeyError as e:
        st.error(f"오류: 파일에 필수 컬럼이 없습니다. 누락된 컬럼: {e}")
        return None
    except Exception as e:
        st.error(f"파일 변환 중 예기치 않은 오류가 발생했습니다. 상세 사유: {e}")
        return None



### **Streamlit 앱 메인 코드**

st.set_page_config(layout="centered") # 페이지 레이아웃 설정
st.title("이카운트 판매입력 변환기")
st.write("이플렉스 주문상세현황 엑셀 파일을 업로드하면 이카운트 '판매입력 웹자료 올리기' 형식으로 변환해 드립니다.")

# 파일 업로더 위젯
uploaded_file = st.file_uploader("쿠팡 주문 엑셀 파일을 선택하세요 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.info(f"파일 '{uploaded_file.name}'을(를) 업로드했습니다. 변환을 시작합니다...")
    
    # 파일 변환 함수 호출
    transformed_df = convert_file(uploaded_file)
    
    if transformed_df is not None:
        st.success("파일 변환이 성공적으로 완료되었습니다!")
        st.dataframe(transformed_df.head()) # 변환된 데이터의 일부를 미리보기로 보여줌

        # 변환된 DataFrame을 엑셀 파일로 메모리에 저장
        output_buffer = BytesIO()
        transformed_df.to_excel(output_buffer, index=False, engine='xlsxwriter')
        output_buffer.seek(0) # 버퍼의 시작 위치로 이동

        # 다운로드 버튼 생성
        st.download_button(
            label="변환된 엑셀 파일 다운로드",
            data=output_buffer,
            file_name="ecount_판매입력_웹자료올리기.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("파일 변환에 실패했습니다. 위의 오류 메시지를 확인해주세요.")

st.markdown("---")
st.caption("개발자: [당신의 이름 또는 회사 이름]") # 선택적으로 개발자 정보 추가