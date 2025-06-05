import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🔍 Customer Validation Tool - KNA1 vs KNVV vs MACE")

# Upload section
st.header("📤 Upload Excel Files")
kna1_file = st.file_uploader("Upload KNA1 Excel", type=["xlsx"])
knvv_file = st.file_uploader("Upload KNVV Excel", type=["xlsx"])
mace_file = st.file_uploader("Upload MACE Excel", type=["xlsx"])

if kna1_file and knvv_file and mace_file:
    if st.button("🔍 Compare"):
        with st.spinner("Processing files and validating customers..."):

            # Read KNA1: header at row 4 (index 4), skip blank row 5 (index 5)
            df_kna1 = pd.read_excel(kna1_file, header=4, skiprows=[5])
            df_kna1.columns = df_kna1.columns.str.strip()
            df_kna1 = df_kna1.loc[:, df_kna1.columns != '']
            if 'Unnamed: 0' in df_kna1.columns:
                df_kna1 = df_kna1.drop(columns=['Unnamed: 0'])
            st.write("KNA1 columns:", df_kna1.columns.tolist())
            st.dataframe(df_kna1.head())

            # Read KNVV: header at row 4 (index 4), skip blank row 5 (index 5)
            df_knvv = pd.read_excel(knvv_file, header=4, skiprows=[5])
            df_knvv.columns = df_knvv.columns.str.strip()
            df_knvv = df_knvv.loc[:, df_knvv.columns != '']
            if 'Unnamed: 0' in df_knvv.columns:
                df_knvv = df_knvv.drop(columns=['Unnamed: 0'])
            st.write("KNVV columns:", df_knvv.columns.tolist())
            st.dataframe(df_knvv.head())

            # Read MACE normally
            df_mace = pd.read_excel(mace_file)
            df_mace.columns = df_mace.columns.str.strip()
            st.write("MACE columns:", df_mace.columns.tolist())
            st.dataframe(df_mace.head())

            # Helper function to find column ignoring case
            def find_column(df, target):
                for col in df.columns:
                    if col.lower() == target.lower():
                        return col
                return None

            kna1_customer_col = find_column(df_kna1, "Customer")
            knvv_customer_col = find_column(df_knvv, "Customer")
            mace_customer_col = find_column(df_mace, "CUSTOMER_NATURAL_ID")

            st.write(f"KNA1 Customer column detected: {kna1_customer_col}")
            st.write(f"KNVV Customer column detected: {knvv_customer_col}")
            st.write(f"MACE Customer column detected: {mace_customer_col}")

            if not (kna1_customer_col and knvv_customer_col and mace_customer_col):
                st.error("One or more required customer ID columns not found!")
                st.stop()

            # Clean data
            df_kna1_clean = df_kna1[df_kna1[kna1_customer_col].notna() & (df_kna1[kna1_customer_col].astype(str).str.strip() != '')]
            df_knvv_clean = df_knvv[df_knvv[knvv_customer_col].notna() & (df_knvv[knvv_customer_col].astype(str).str.strip() != '')]
            df_mace_clean = df_mace[df_mace[mace_customer_col].notna() & (df_mace[mace_customer_col].astype(str).str.strip() != '')]

            st.write(f"KNA1 rows before: {len(df_kna1)}, after cleaning: {len(df_kna1_clean)}")
            st.write(f"KNVV rows before: {len(df_knvv)}, after cleaning: {len(df_knvv_clean)}")
            st.write(f"MACE rows before: {len(df_mace)}, after cleaning: {len(df_mace_clean)}")

            st.write("KNA1 cleaned preview:")
            st.dataframe(df_kna1_clean.head())
            st.write("KNVV cleaned preview:")
            st.dataframe(df_knvv_clean.head())
            st.write("MACE cleaned preview:")
            st.dataframe(df_mace_clean.head())

            # Customer sets
            kna1_customers = set(df_kna1_clean[kna1_customer_col].astype(str).str.strip())
            knvv_customers = set(df_knvv_clean[knvv_customer_col].astype(str).str.strip())
            mace_customers = set(df_mace_clean[mace_customer_col].astype(str).str.strip())

            df_diff1 = pd.DataFrame(sorted(kna1_customers - knvv_customers), columns=["Customer_Not_in_KNVV"])
            df_diff2 = pd.DataFrame(sorted(knvv_customers - mace_customers), columns=["Customer_Not_in_MACE"])

            # Start index from 1 in UI
            df_diff1_display = df_diff1.reset_index(drop=True)
            df_diff1_display.index += 1
            df_diff2_display = df_diff2.reset_index(drop=True)
            df_diff2_display.index += 1

            st.subheader("❗ Customers in KNA1 but NOT in KNVV")
            st.dataframe(df_diff1_display, use_container_width=True)

            st.subheader("❗ Customers in KNVV but NOT in MACE")
            st.dataframe(df_diff2_display, use_container_width=True)

            # Excel export
            def to_excel(df1, df2):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df1.to_excel(writer, index=False, sheet_name='KNA1_Not_in_KNVV')
                    df2.to_excel(writer, index=False, sheet_name='KNVV_Not_in_MACE')
                output.seek(0)
                return output

            excel_data = to_excel(df_diff1, df_diff2)
            st.download_button("⬇️ Download Validation Result", excel_data, file_name="customer_validation_result.xlsx")

else:
    st.warning("Please upload all 3 files to proceed.")
