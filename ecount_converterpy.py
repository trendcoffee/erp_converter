
import streamlit as st
import pandas as pd
import math
import io

st.set_page_config(page_title="이카운트 변환기", layout="centered")
st.title("이카운트 판매입력 자동 변환기")

uploaded_file = st.file_uploader("이플렉스 주문상세현황 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # 금액과 주문수량 정리
        df["금액"] = df["금액"].astype(str).str.replace(",", "")
        df["금액"] = pd.to_numeric(df["금액"], errors="coerce")
        df["주문수량"] = pd.to_numeric(df["주문수량"], errors="coerce")

        df = df.dropna(subset=["금액", "주문수량"])
        df = df[df["주문수량"] != 0]

        단가 = (df["금액"] / df["주문수량"]).fillna(0)
        수량 = df["주문수량"].fillna(0)
        총금액 = (단가 * 수량).round().fillna(0)
        부가세 = (총금액 / 11).fillna(0).astype(int)
        공급가액 = (총금액 - 부가세).fillna(0).astype(int)

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

        거래처명, 수집처 = zip(*df["판매채널"].map(get_client_name))

        생산전표_N_목록 = [
            "YELLOW_NOZZLE_2EA", "WHITE_NOZZLE_2EA", "RED_NOZZLE_2EA",
            "YELLOW_T_NOZZLE-CAP", "WHITE_T_NOZZLE-CAP"
        ]

        result = pd.DataFrame()
        result["일자"] = pd.to_datetime(df["출고일자"]).dt.strftime("%Y%m%d")
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
        result["운송장번호"] = df["대표운송장번호"].astype(str).str.replace(".0", "", regex=False)
        result["적요"] = ""
        result["생산전표생성"] = result["품목코드"].apply(lambda x: "N" if str(x).strip() in 생산전표_N_목록 else "Y")

        columns_order = [
            "일자", "순번", "거래처코드", "거래처명", "담당자", "출하창고", "거래유형", "통화", "환율", "잔액", "참고",
            "품목코드", "품목명", "규격", "수량", "단가", "외화금액", "공급가액", "부가세",
            "수집처", "수취인", "운송장번호", "적요", "생산전표생성"
        ]
        result = result[columns_order]
        result = result[:-1]

        output = io.BytesIO()
        result.to_excel(output, index=False)
        st.success("✅ 변환 완료! 아래 버튼으로 다운로드하세요.")
        st.download_button("📥 변환된 파일 다운로드", output.getvalue(), file_name="ecount_output.xlsx")

    except Exception as e:
        st.error(f"❌ 오류 발생: {str(e)}")
