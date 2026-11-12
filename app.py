import streamlit as st
import pandas as pd
import io

# 1. 페이지 기본 설정 (와이드 레이아웃 사용)
st.set_page_config(
    page_title="생산 관리 자동화 시스템",
    page_icon="⚙️",
    layout="wide"
)

# --- [사이드바 영역] 꾸미기 및 메뉴 구성 ---
st.sidebar.title("🛠️ 생산관리 시스템")
st.sidebar.markdown("---")

# 메뉴 선택 라디오 버튼
menu = st.sidebar.radio(
    "메뉴 선택",
    ["📁 주간 사용 리미트캠", "📊 현황 대시보드 (준비중)", "📖 사용 가이드"]
)

st.sidebar.markdown("---")
st.sidebar.info("💡 문의 사항이나 기능 추가 요청은 박희준과장에게 문의하세요.")


# --- [메인 화면 1] 주간 사용 리미트캠 (기존 기능) ---
if menu == "📁 주간 사용 리미트캠":
    st.title("📁 주간 사용 리미트캠")
    st.markdown("엑셀 파일을 업로드하면 **전체 데이터**와 **추출 데이터**를 비교하여 양식에 맞게 자동 정리해 줍니다.")
    
    # 탭으로 화면 안에서 구역 나누기
    tab1, tab2 = st.tabs(["📤 파일 업로드 및 실행", "ℹ️ 이용 안내"])
    
    with tab1:
        uploaded_file = st.file_uploader("엑셀 파일을 선택하거나 여기에 드래그하세요 (.xlsx)", type=["xlsx"])

        if uploaded_file is not None:
            try:
                xls = pd.ExcelFile(uploaded_file)
                sheet_names = xls.sheet_names
                
                # 가로로 컬럼을 나누어 시트 선택 배치
                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    default_all = sheet_names.index("전체데이터") if "전체데이터" in sheet_names else 0
                    sheet_all_name = st.selectbox("전체 데이터 시트 선택", sheet_names, index=default_all)
                with col_s2:
                    default_ext = sheet_names.index("추출데이터") if "추출데이터" in sheet_names else (1 if len(sheet_names) > 1 else 0)
                    sheet_ext_name = st.selectbox("추출 기준 시트 (제번 목록) 선택", sheet_names, index=default_ext)
                
                df_all = pd.read_excel(uploaded_file, sheet_name=sheet_all_name)
                df_extract = pd.read_excel(uploaded_file, sheet_name=sheet_ext_name)
                
                ext_col = df_extract.columns[0]
                extract_list = df_extract[ext_col].dropna().astype(str).str.strip().unique().tolist()
                
                if "제번" not in df_all.columns:
                    st.error("선택한 '전체 데이터 시트'에 '제번' 컬럼이 없습니다.")
                else:
                    df_all["제번_정제"] = df_all["제번"].astype(str).str.strip()
                    filtered_df = df_all[df_all["제번_정제"].isin(extract_list)].copy()
                    
                    group_cols = ["제품PT#", "품 명", "규 격"]
                    if all(col in filtered_df.columns for col in group_cols):
                        result_df = (
                            filtered_df.groupby(group_cols, dropna=False)
                            .size()
                            .reset_index(name="개수 : 제품PT#")
                        )
                        
                        total_count = result_df["개수 : 제품PT#"].sum()
                        
                        total_row = pd.DataFrame([{
                            "제품PT#": "",
                            "품 명": "",
                            "규 격": "총합계",
                            "개수 : 제품PT#": f"{total_count}EA"
                        }])
                        final_df = pd.concat([result_df, total_row], ignore_index=True)
                        
                        st.markdown("---")
                        st.success(f"✨ 처리 완료! (총 제번: {len(extract_list)}건 중 {len(filtered_df)}건 매칭)")
                        
                        # 화면 분할 배치 (표와 요약 지표)
                        res_col1, res_col2 = st.columns([3, 1])
                        with res_col1:
                            st.subheader("📋 정리된 결과물")
                            st.dataframe(final_df, use_container_width=True, height=400)
                        
                        with res_col2:
                            st.subheader("📊 요약 정보")
                            st.metric(label="총 수량", value=f"{total_count} EA")
                            st.metric(label="품목 종류", value=f"{len(result_df)} 종")
                            
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                                final_df.to_excel(writer, index=False, sheet_name="결과물")
                                filtered_df.drop(columns=["제번_정제"]).to_excel(writer, index=False, sheet_name="전체에서추출")
                            
                            st.download_button(
                                label="📥 결과 엑셀 다운로드",
                                data=output.getvalue(),
                                file_name="자동정리_결과물.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                    else:
                        st.error(f"필수 컬럼({group_cols})이 존재하지 않습니다.")
            except Exception as e:
                st.error(f"파일 처리 중 오류 발생: {e}")
        else:
            st.info("👆 위 영역에 엑셀 파일을 업로드해 주세요.")
            
    with tab2:
        st.markdown("""
        ### 📌 사용 방법 요약
        1. 업로드할 엑셀 파일에 **전체데이터** 시트와 **추출데이터** 시트가 포함되어 있어야 합니다.
        2. 파일 업로드 후 시트 이름이 올바르게 선택되었는지 확인하세요.
        3. 결과가 화면에 나오면 우측의 **[다운로드 버튼]**을 눌러 엑셀 파일을 저장할 수 있습니다.
        """)

# --- [메인 화면 2] 현황 대시보드 메뉴 ---
elif menu == "📊 현황 대시보드 (준비중)":
    st.title("📊 생산 현황 대시보드")
    st.info("차후 생산 데이터 통계 및 그래프를 시각화해서 보여줄 공간입니다.")

# --- [메인 화면 3] 사용 가이드 메뉴 ---
elif menu == "📖 사용 가이드":
    st.title("📖 시스템 사용 가이드")
    st.markdown("""
    ### 🚀 프로그램 소개
    본 웹 프로그램은 수작업으로 처리하던 생산 부품 제번 추출 및 집계 업무를 웹 브라우저에서 간편하게 처리하기 위해 제작되었습니다.
    
    - **개발 환경:** Python + Streamlit
    - **배포 서버:** Streamlit Community Cloud
    """)
