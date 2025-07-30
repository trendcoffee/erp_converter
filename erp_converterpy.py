import pandas as pd
import os
import streamlit as st
from io import BytesIO

def convert_file(uploaded_file):
    """
    업로드된 엑셀 파일을 받아 이카운트 형식으로 변환하는 함수.
    Args:
        uploaded_file: Streamlit의 st.file_uploader를 통해 업로드된 파일 객체.
    Returns:
        변환된 DataFrame (성공 시) 또는 None (실패 시).
    """
    try:
        df = pd.read_excel(uploaded_file)
        df.fillna('', inplace=True) 

        df["금액"] = df["금액"].astype(str).str.replace(",", "")
        df["금액"] = pd.to_numeric(df["금액"], errors="coerce").fillna(0)
        
        df["주문수량"] = pd.to_numeric(df["주문수량"], errors="coerce").fillna(0)
        
        df["대표운송장번호"] = pd.to_numeric(df["대표운송장번호"], errors="coerce").astype(pd.Int64Dtype())
        
        df = df[df["주문수량"] != 0]

        단가 = (df["금액"] / df["주문수량"]).fillna(0) 
        수량 = df["주문수량"].fillna(0)
        총금액 = (단가 * 수량).round().fillna(0)
        부가세 = (총금액 / 11).fillna(0).astype(int)
        공급가액 = (총금액 - 부가세).fillna(0).astype(int)

        # --- get_client_name 함수 수정 시작 ---
        # 이 함수는 이제 '판매채널'과 '주문중개채널(상세)' 두 값을 받습니다.
        def get_client_name(row):
            판매채널_val = str(row["판매채널"]).strip().upper()
            주문중개채널_상세_val = str(row["주문중개채널(상세)"]).strip().upper() if "주문중개채널(상세)" in row else ""

            # '주문중개채널(상세)'가 '쿠팡' 또는 'SSG'일 경우 이를 우선 적용
            if 주문중개채널_상세_val == "COUPANG":
                return "쿠팡 주식회사", "쿠팡"
            elif 주문중개채널_상세_val == "SSG":
                return "(주)에스에스지닷컴", "SSG"
            # 그 외의 경우 기존 '판매채널' 로직 유지
            elif 판매채널_val == "GMKT":
                return "지마켓글로벌 유한책임회사", "지마켓"
            elif 판매채널_val == "AUCT":
                return "지마켓글로벌 유한책임회사", "옥션"
            elif 판매채널_val == "롯데ON":
                return "롯데쇼핑주식회사", "롯데온"
            elif 판매채널_val == "NFA": # 네이버페이 (스마트스토어)
                return "네이버파이낸셜 주식회사", "스마트스토어"
            elif 판매채널_val == "11번가":
                return "십일번가 주식회사", "11번가"
            else: # 매핑되지 않은 경우 판매채널 값을 그대로 사용
                return 판매채널_val, 판매채널_val
        # --- get_client_name 함수 수정 끝 ---

        # '판매채널' 및 '주문중개채널(상세)' 컬럼 존재 여부 확인 후 처리
        required_input_cols = ["판매채널"]
        if "주문중개채널(상세)" not in df.columns:
            st.warning("경고: '주문중개채널(상세)' 컬럼이 파일에 없습니다. '판매채널' 정보만 사용하여 처리합니다.")
        # 모든 필수 컬럼이 있는지 다시 한 번 확인
        for col in required_input_cols:
            if col not in df.columns:
                st.error(f"오류: 필수 컬럼 '{col}'이(가) 파일에 없습니다.")
                return None
            
        # apply 함수에 axis=1을 사용하여 각 행(row)을 get_client_name에 전달
        # 반환된 튜플을 두 개의 리스트로 언팩(unpack)
        거래처명, 수집처 = zip(*df.apply(get_client_name, axis=1))


        생산전표_N_목록 = [
            "YELLOW_NOZZLE_2EA", "WHITE_NOZZLE_2EA", "RED_NOZZLE_2EA",
            "YELLOW_T_NOZZLE-CAP", "WHITE_T_NOZZLE-CAP"
        ]

        result = pd.DataFrame()
        if "출고일자" not in df.columns:
            st.error("오류: 필수 컬럼 '출고일자'이(가) 파일에 없습니다.")
            return None
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
        result["운송장번호"] = df["대표운송장번호"].astype(str).replace('<NA>', '')
        result["적요"] = ""
        result["생산전표생성"] = result["품목코드"].apply(lambda x: "N" if str(x).strip() in 생산전표_N_목록 else "Y")

        columns_order = [
            "일자", "순번", "거래처코드", "거래처명", "담당자", "출하창고", "거래유형", "통화", "환율", "잔액", "참고",
            "품목코드", "품목명", "규격", "수량", "단가", "외화금액", "공급가액", "부가세",
            "수집처", "수취인", "운송장번호", "적요", "생산전표생성"
        ]
        result = result[columns_order]
        
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

st.set_page_config(layout="centered")
st.title("판매 입력 변환기")
st.write("이플렉스 주문상세현황 엑셀 파일을 업로드하면 이카운트 '판매입력 웹자료 올리기' 형식으로 변환합니다.")

uploaded_file = st.file_uploader("이플렉스 주문상세현황 엑셀 파일을 선택하세요 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.info(f"파일 '{uploaded_file.name}'을(를) 업로드했습니다. 변환을 시작합니다...")
    
    transformed_df = convert_file(uploaded_file)
    
    if transformed_df is not None:
        st.success("파일 변환이 성공적으로 완료되었습니다!")
        st.dataframe(transformed_df.head())

        output_buffer = BytesIO()
        transformed_df.to_excel(output_buffer, index=False, engine='xlsxwriter')
        output_buffer.seek(0)

        st.download_button(
            label="변환된 엑셀 파일 다운로드",
            data=output_buffer,
            file_name="ecount_판매입력_웹자료올리기.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("파일 변환에 실패했습니다. 위의 오류 메시지를 확인해주세요.")

st.markdown("---")
st.caption("[더끌림컴퍼니]")