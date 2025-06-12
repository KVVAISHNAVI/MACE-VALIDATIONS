import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("🔍 Customer Validation Tool")

tabs = st.tabs(["📄 KNA1 vs KNVV", "📄 KNA1+KNVV vs MACE"])

# ---------- TAB 1: KNA1 vs KNVV ----------
with tabs[0]:
    st.header("📤 Upload KNA1 and KNVV Files")
    kna1_file = st.file_uploader("Upload KNA1 Excel", type=["xlsx"], key="kna1")
    knvv_file = st.file_uploader("Upload KNVV Excel", type=["xlsx"], key="knvv")

    def find_column(df, target):
        for col in df.columns:
            if col.lower() == target.lower():
                return col
        return None

    def clean_all_text_columns(df):
        df = df.fillna('').replace({pd.NA: ''})
        for col in df.columns:
            df[col] = df[col].astype(str).str.replace(r"\s+", " ", regex=True).str.replace("\xa0", " ", regex=True).str.strip()
        return df

    if kna1_file and knvv_file:
        if st.button("🔍 Compare", key="compare_kna1_knvv"):
            with st.spinner("Validating..."):
                # Read and clean KNA1
                df_kna1 = pd.read_excel(kna1_file, header=4, skiprows=[5])
                df_kna1.columns = df_kna1.columns.str.strip()
                df_kna1 = df_kna1.loc[:, ~df_kna1.columns.str.contains('^Unnamed', case=False) & (df_kna1.columns.str.strip() != '')]
                df_kna1 = clean_all_text_columns(df_kna1)

                # Read and clean KNVV
                df_knvv = pd.read_excel(knvv_file, header=4, skiprows=[5])
                df_knvv.columns = df_knvv.columns.str.strip()
                df_knvv = df_knvv.loc[:, ~df_knvv.columns.str.contains('^Unnamed', case=False) & (df_knvv.columns.str.strip() != '')]
                df_knvv = clean_all_text_columns(df_knvv)

                # Find "Customer" column
                customer_col_kna1 = find_column(df_kna1, "Customer")
                customer_col_knvv = find_column(df_knvv, "Customer")

                # Filter non-empty customers
                df_kna1_clean = df_kna1[df_kna1[customer_col_kna1] != '']
                df_knvv_clean = df_knvv[df_knvv[customer_col_knvv] != '']

                # Unique customers
                kna1_customers = set(df_kna1_clean[customer_col_kna1])
                knvv_customers = set(df_knvv_clean[customer_col_knvv])

                # Differences
                df_diff1 = df_kna1_clean[df_kna1_clean[customer_col_kna1].isin(kna1_customers - knvv_customers)]
                df_diff2 = df_knvv_clean[df_knvv_clean[customer_col_knvv].isin(knvv_customers - kna1_customers)]
                df_diff1.index = range(1, len(df_diff1) + 1)
                df_diff2.index = range(1, len(df_diff2) + 1)

                st.write(f"🔢 Customers in KNA1 not in KNVV: {len(df_diff1)}")
                st.write(f"🔢 Customers in KNVV not in KNA1: {len(df_diff2)}")

                st.subheader("❗ Customers in KNA1 but NOT in KNVV")
                st.dataframe(df_diff1)

                st.subheader("❗ Customers in KNVV but NOT in KNA1")
                st.dataframe(df_diff2)

                # Merge on Customer only
                merged_df = pd.merge(
                    df_kna1_clean,
                    df_knvv_clean,
                    how="left",
                    left_on=customer_col_kna1,
                    right_on=customer_col_knvv,
                    suffixes=('', '_KNVV')
                )

                # Remove duplicate columns from KNVV (_KNVV)
                merged_df = merged_df.drop(columns=[col for col in merged_df.columns if col.endswith('_KNVV')])

                st.subheader("🔗 Merged View")
                merged_df.index = range(1, len(merged_df) + 1)
                st.dataframe(merged_df)

                # Excel download function
                def to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    output.seek(0)
                    return output

                st.download_button("⬇️ Download Merged Excel", to_excel(merged_df), file_name="merged_kna1_knvv.xlsx")

# ---------- TAB 2: Merged vs MACE ----------
with tabs[1]:
    st.header("📥 Upload KNA1+KNVV and MACE File")
    merged_file = st.file_uploader("Upload KNA1+KNVV Excel from Tab 1", type=["xlsx"], key="merged")
    mace_file = st.file_uploader("Upload MACE Excel", type=["xlsx"], key="mace")

    if merged_file and mace_file:
        if st.button("📎 Compare with MACE", key="compare_mace"):
            with st.spinner("Comparing KNA1+KNVV data with MACE..."):
                try:
                    df_merged = pd.read_excel(merged_file)
                    df_mace = pd.read_excel(mace_file)
                    df_merged = clean_all_text_columns(df_merged)
                    df_mace = clean_all_text_columns(df_mace)
                except Exception as e:
                    st.error(f"Error reading files: {e}")
                    st.stop()

                if "CUSTOMER_NATURAL_ID" not in df_mace.columns:
                    st.error("❌ MACE file must contain 'CUSTOMER_NATURAL_ID'")
                    st.stop()

                # Define column mapping
                column_mapping = {
                    "Customer": "CUSTOMER_NATURAL_ID",
                    "City": "CUSTOMER_CITY_NAME",
                    "Ctry/Reg.": "CUSTOMER_COUNTRY_ISO2_CODE",
                    "Postal Code": "CUSTOMER_POSTAL_CODE",
                    "Street": "CUSTOMER_STREET_NAME",
                    "Region": "CUSTOMER_REGION_CODE",
                    "Name": "CUSTOMER_NAME",
                    "Name2": "CUSTOMER_NAME2",
                    "Sales Org.": "CUSTOMER_SALES_ORGANIZATION_CODE",
                    "Distr. Channel": "CUSTOMER_SALES_DISTRIBUTION_CHANNEL_CODE",
                    "Division": "CUSTOMER_DIVISION_CODE",
                    "Currency": "CUSTOMER_CURRENCY",
                    "Account group": "CUSTOMER_ACCOUNT_GROUP_CODE",
                    "Language": "CUSTOMER_LANGUAGE_KEY",
                    "Group": "CUSTOMER_GROUP_KEY"
                }

                merged_not_in_mace = []
                mismatch_reason = []
                mace_customers = set(df_mace["CUSTOMER_NATURAL_ID"].astype(str).str.strip())
                merged_customers = df_merged["Customer"].astype(str).str.strip()

                for idx, row in df_merged.iterrows():
                    cust_id = str(row.get("Customer", "")).strip()

                    # If customer is not in MACE at all
                    if cust_id not in mace_customers:
                        merged_not_in_mace.append(row)
                        mismatch_reason.append("Customer not found in MACE")
                        continue

                    # Check all MACE rows for this customer
                    matching_mace_rows = df_mace[df_mace["CUSTOMER_NATURAL_ID"] == cust_id]
                    found_match = False
                    mismatch_cols = []

                    for _, mace_row in matching_mace_rows.iterrows():
                        current_mismatch = []

                        for m_col, mace_col in column_mapping.items():
                            if m_col not in row or mace_col not in mace_row:
                                continue

                            val_merged = str(row[m_col]).strip()
                            val_mace = str(mace_row[mace_col]).strip()

                            # Skip blanks or 'not found'
                            if (
                                val_merged == ""
                                or val_mace == ""
                                or val_merged.lower() == "not found"
                                or val_mace.lower() == "not found"
                            ):
                                continue

                            try:
                                if float(val_merged) != float(val_mace):
                                    current_mismatch.append(m_col)
                            except:
                                if val_merged != val_mace:
                                    current_mismatch.append(m_col)

                        # If exact match found
                        if not current_mismatch:
                            found_match = True
                            break

                        # Track first mismatch set
                        if not mismatch_cols:
                            mismatch_cols = current_mismatch

                    if not found_match:
                        merged_not_in_mace.append(row)
                        mismatch_reason.append(", ".join(mismatch_cols) if mismatch_cols else "Mismatch")

                df_not_in_mace = pd.DataFrame(merged_not_in_mace)
                if not df_not_in_mace.empty:
                    df_not_in_mace["Mismatch Reason"] = mismatch_reason
                    df_not_in_mace.index = range(1, len(df_not_in_mace) + 1)

                # Customers in MACE not in merged
                mace_customers_set = set(df_mace["CUSTOMER_NATURAL_ID"].astype(str).str.strip())
                merged_customers_set = set(merged_customers)
                df_not_in_merged = df_mace[df_mace["CUSTOMER_NATURAL_ID"].isin(mace_customers_set - merged_customers_set)]
                df_not_in_merged.index = range(1, len(df_not_in_merged) + 1)

                # Show results
                st.write(f"🔢 Customers in KNA1+KNVV but NOT in MACE (including mismatches): {len(df_not_in_mace)}")
                st.write(f"🔢 Customers in MACE but NOT in KNA1+KNVV: {len(df_not_in_merged)}")

                st.subheader("❗ Customers in KNA1+KNVV but NOT in MACE (or mismatched)")
                st.dataframe(df_not_in_mace)

                st.subheader("❗ Customers in MACE but NOT in KNA1+KNVV")
                st.dataframe(df_not_in_merged)

                # Download button
                def download_excel(df1, df2):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df1.to_excel(writer, index=False, sheet_name='kna1+knvv _Not_in_MACE')
                        df2.to_excel(writer, index=False, sheet_name='MACE_Not_in_ kna1+knvv')
                    output.seek(0)
                    return output

                st.download_button("⬇️ Download MACE_kna1+knvv Comparison Result", download_excel(df_not_in_mace, df_not_in_merged), file_name="mace_kna1+knvv_comparison.xlsx")
