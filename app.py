import streamlit as st
import pandas as pd
import io

# 페이지 기본 설정
st.set_page_config(page_title="생산 데이터 자동 정리 시스템", layout="wide")

st.title("📊 제번 기준 데이터 자동 추출 및 집계 프로그램")
st.markdown("엑셀 파일(`전체데이터` 및 `추출데이터` 시트 포함)을 업로드하면 자동으로 양식에 맞춰 정리하고 다운로드할 수 있습니다.")

# 1. 파일 업로더 위젯
uploaded_file = st.file_uploader("엑셀 파일을 선택하거나 여기로 드래그하세요 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 엑셀 시트 목록 읽기
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        st.sidebar.header("⚙️ 시트 설정")
        # 기본 시트 이름 자동 선택 (없으면 첫 번째/두 번째 시트 선택)
        default_all = sheet_names.index("전체데이터") if "전체데이터" in sheet_names else 0
        default_ext = sheet_names.index("추출데이터") if "추출데이터" in sheet_names else (1 if len(sheet_names) > 1 else 0)
        
        sheet_all_name = st.sidebar.selectbox("전체 데이터 시트", sheet_names, index=default_all)
        sheet_ext_name = st.sidebar.selectbox("추출 기준 시트 (제번 목록)", sheet_names, index=default_ext)
        
        # 데이터 불러오기
        df_all = pd.read_excel(uploaded_file, sheet_name=sheet_all_name)
        df_extract = pd.read_excel(uploaded_file, sheet_name=sheet_ext_name)
        
        # 2. 데이터 전처리 및 필터링
        # 추출데이터의 첫 번째 열(제번) 공백 제거
        ext_col = df_extract.columns[0]
        extract_list = df_extract[ext_col].dropna().astype(str).str.strip().unique().tolist()
        
        # 전체데이터의 '제번' 열 공백 제거 후 필터링
        if "제번" not in df_all.columns:
            st.error("선택한 '전체 데이터 시트'에 '제번' 컬럼이 없습니다. 시트 설정을 확인해 주세요.")
        else:
            df_all["제번_정제"] = df_all["제번"].astype(str).str.strip()
            filtered_df = df_all[df_all["제번_정제"].isin(extract_list)].copy()
            
            # 3. 제품PT#, 품 명, 규 격 기준 그룹화 및 개수 집계
            group_cols = ["제품PT#", "품 명", "규 격"]
            if all(col in filtered_df.columns for col in group_cols):
                result_df = (
                    filtered_df.groupby(group_cols, dropna=False)
                    .size()
                    .reset_index(name="개수 : 제품PT#")
                )
                
                total_count = result_df["개수 : 제품PT#"].sum()
                
                # 총합계 행 추가
                total_row = pd.DataFrame([{
                    "제품PT#": "",
                    "품 명": "",
                    "규 격": "총합계",
                    "개수 : 제품PT#": f"{total_count}EA"
                }])
                final_df = pd.concat([result_df, total_row], ignore_index=True)
                
                # 4. 웹 화면에 결과 출력
                st.success(f"처리 완료! 총 {len(extract_list)}개의 기준 제번 중 {len(filtered_df)}건의 데이터가 추출되었습니다.")
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.subheader("📋 정리된 결과물 미리보기")
                    st.dataframe(final_df, use_container_width=True, height=500)
                
                with col2:
                    st.subheader("🔍 추출된 원본 내역 요약")
                    st.metric(label="총 추출 수량", value=f"{total_count} EA")
                    st.metric(label="품목 종류 수", value=f"{len(result_df)} 종")
                    
                    # 5. 엑셀 파일 다운로드 버튼 생성
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        final_df.to_excel(writer, index=False, sheet_name="결과물")
                        filtered_df.drop(columns=["제번_정제"]).to_excel(writer, index=False, sheet_name="전체에서추출")
                    
                    st.download_button(
                        label="📥 결과물 엑셀 파일 다운로드",
                        data=output.getvalue(),
                        file_name="자동정리_결과물.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
            else:
                st.error(f"필수 컬럼({group_cols})이 전체데이터 시트에 존재하지 않습니다.")
                
    except Exception as e:
        st.error(f"파일 처리 중 오류가 발생했습니다: {e}")
else:
    st.info("👆 위 버튼을 눌러 작업할 엑셀 파일을 업로드해 주세요.")