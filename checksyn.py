import pandas as pd

# ======================
# Stage 1: NO PO CHECK
# ======================
wms_file = "WMS.xlsx"
sap_file = "SAP.xlsx"
wms_file_po = "036.xls"


df_wms = pd.read_excel(wms_file)
df_sap = pd.read_excel(sap_file, header=2)

# Ensure numeric conversion for WMS (columns C→H)
cols_to_fix = df_wms.columns[2:8]  # adjust if needed
df_wms[cols_to_fix] = (
    df_wms[cols_to_fix]
    .replace(r"[^\d.-]", "", regex=True)
    .astype(float)
    .fillna(0)
    .astype(int)
)

# Calculate WMS values
df_wms["UU (WMS)"] = df_wms.iloc[:, 3] + df_wms.iloc[:, 4] + df_wms.iloc[:, 5]   # D+E+F
df_wms["BLOCK (WMS)"] = df_wms.iloc[:, 6]                                        # G
df_wms["TOTAL QTY (WMS)"] = df_wms.iloc[:, 2]                                    # C

# Group SAP by Material
sap_grouped = df_sap.groupby("Material").agg({
    "Unrestricted Use Qty": "sum",
    "Block Stock": "sum",
    "Total(UU+QI+Blocked)": "sum"
}).reset_index()

sap_grouped.rename(columns={
    "Unrestricted Use Qty": "UU (SAP)",
    "Block Stock": "BLOCK (SAP)",
    "Total(UU+QI+Blocked)": "TOTAL QTY (SAP)"
}, inplace=True)

# Merge SAP into WMS (no PO check)
df_result = df_wms.merge(
    sap_grouped, how="left", 
    left_on=df_wms.columns[0], right_on="Material"
)

df_result[["UU (SAP)", "BLOCK (SAP)", "TOTAL QTY (SAP)"]] = (
    df_result[["UU (SAP)", "BLOCK (SAP)", "TOTAL QTY (SAP)"]].fillna(0).astype(int)
)

# Calculate differences
df_result["LỆCH BTP (WMS - SAP)"] = df_result["UU (WMS)"] - df_result["UU (SAP)"]
df_result["LỆCH HOLD (WMS - SAP)"] = df_result["BLOCK (WMS)"] - df_result["BLOCK (SAP)"]
df_result["LỆCH TỔNG TỒN (WMS - SAP)"] = df_result["TOTAL QTY (WMS)"] - df_result["TOTAL QTY (SAP)"]

# Final columns for Stage 1
final_columns = [
    df_wms.columns[0],
    "UU (WMS)", "BLOCK (WMS)", "UU (SAP)", "BLOCK (SAP)",
    "TOTAL QTY (WMS)", "TOTAL QTY (SAP)",
    "LỆCH BTP (WMS - SAP)", "LỆCH HOLD (WMS - SAP)", "LỆCH TỔNG TỒN (WMS - SAP)"
]
df_final = df_result[final_columns]

# ======================
# Stage 2: WITH PO CHECK
# ======================

df_wms_po = pd.read_excel(wms_file_po, header=3)  # skip 3 rows, 4th row header

# Rename columns sequentially
df_wms_po.columns = [
    "NO", "LOC", "ITEM", "NAMEITEM", "QTY", "QTYS", 
    "LPN", "PO", "NCC", "ReceiptDate", "OrderDate"
]

# Drop blanks in ITEM
df_wms_po = df_wms_po[df_wms_po["ITEM"].notna() & (df_wms_po["ITEM"].astype(str).str.strip() != "")]

# Aggregate WMS by ITEM + PO
wms_grouped_po = df_wms_po.groupby(["ITEM", "PO"]).agg({
    "QTY": "sum"
}).reset_index().rename(columns={"QTY": "QTY BLOCK WMS"})

# Aggregate SAP by Material + PO
sap_grouped_po = df_sap.groupby(["Material", "PO MPE"]).agg({
    "Block Stock": "sum"
}).reset_index().rename(columns={
    "Material": "ITEM",
    "PO MPE": "PO",
    "Block Stock": "QTY BLOCK SAP"
})

# Merge both
sync_po = pd.merge(wms_grouped_po, sap_grouped_po, on=["ITEM", "PO"], how="outer").fillna(0)
sync_po["QTY LECH BLOCK WMS - SAP"] = sync_po["QTY BLOCK WMS"] - sync_po["QTY BLOCK SAP"]

# ======================
# Export Both Sheets
# ======================
with pd.ExcelWriter("Check_Sync_WMS_vs_SAP.xlsx", engine="openpyxl") as writer:
    df_final.to_excel(writer, sheet_name="NO_PO", index=False)   # Stage 1
    sync_po.to_excel(writer, sheet_name="WITH_PO", index=False)  # Stage 2
