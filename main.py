import pandas as pd

# =============================
# Config
# =============================
FILE_NAME = "input/HLHL.xlsx"
DATASET_NAME = "HLHL"

SUMMARY_PRODUCTS = [
    "Total Capacity",
    "Hoplite (FBG) 980",
    "Hoplite (FBG) 14xx",
    "Hoplens (FTA) 980",
    "Hoplens (FLA) 14xx",
    "Hoplens (FTA) H1x",
    "(FBG/FLA) 14xx Turn-Key",
    "Hoplite",
    "Hoplens",
    "Total HLHL",
    "Hoplite (FBG) HRS",
    "Hoplens (FTA) HRS",
    "HRS"
]

TARGET_PRODUCTS = [
    "Total build plan",
    "Hoplite (FBG) 980",
    "Hoplite (FBG) 14xx",
    "Hoplens (FTA) 980",
    "Hoplens (FLA) 14xx",
    "Hoplens (FTA) H1x",
    "(FBG/FLA) 14xx Turn-Key",
    "Hoplite",
    "Hoplens",
    "Total HLHL",
    "Hoplite (FBG) HRS",
    "Hoplens (FTA) HRS",
    "HRS"
]

# Weekly summary rows
CAPACITY_WEEKLY_ROW_START = 38   # row 39
CAPACITY_WEEKLY_ROW_END = 51     # row 51
DEMAND_WEEKLY_ROW_START = 9      # row 10
DEMAND_WEEKLY_ROW_END = 22       # row 22

# Quarterly summary rows
CAPACITY_QUARTERLY_ROW_START = 23   # row 24
CAPACITY_QUARTERLY_ROW_END = 36     # row 36
DEMAND_QUARTERLY_ROW_START = 9      # row 10
DEMAND_QUARTERLY_ROW_END = 22       # row 22

# Target rows
TARGET_WEEKLY_ROW_START = 23   # row 24
TARGET_WEEKLY_ROW_END = 36     # row 36
TARGET_QUARTERLY_ROW_START = 23
TARGET_QUARTERLY_ROW_END = 36

# Weekly blocks
WEEKLY_BLOCKS = [
    {"week_start": 5,  "week_end": 18, "quarter_col": 5},   # F:R
    {"week_start": 19, "week_end": 32, "quarter_col": 19},  # T:AF
    {"week_start": 33, "week_end": 46, "quarter_col": 33},  # AH:AT
    {"week_start": 47, "week_end": 60, "quarter_col": 47},  # AV:BH
]

# Quarterly columns
QUARTERLY_COLS = [
    {"label": "Q326", "col": 18},  # S
    {"label": "Q426", "col": 32},  # AG
    {"label": "Q127", "col": 46},  # AU
    {"label": "Q227", "col": 60},  # BI
]

# MPU blocks by product
MPU_PRODUCT_BLOCKS = [
    {"product": "Hoplite (FBG) 980",              "row_start": 127, "row_end": 137},  # 128:137
    {"product": "Hoplite (FBG) 14xx",             "row_start": 137, "row_end": 147},  # 138:147
    {"product": "Hoplite (FBG) HRS",              "row_start": 147, "row_end": 157},  # 148:157
    {"product": "Hoplens (FTA) 980",              "row_start": 157, "row_end": 166},  # 158:166
    {"product": "Hoplens (FLA) 14xx",             "row_start": 166, "row_end": 175},  # 167:175
    {"product": "Hoplens (FTA) H1x",              "row_start": 175, "row_end": 185},  # 176:185
    {"product": "Hoplite (FBG) 14xx Turn-Key",    "row_start": 185, "row_end": 195},  # 186:195
    {"product": "Hoplens (FLA) 14xx Turn-Key",    "row_start": 195, "row_end": 204},  # 196:204
    {"product": "Hoplens (FTA) HRS",              "row_start": 204, "row_end": 213},  # 205:213
]

OUTPUT_FILE = "output/capacity_dashboard_dataset.xlsx"


# =============================
# Helper functions
# =============================
def extract_weekly_block(
    df: pd.DataFrame,
    row_start: int,
    row_end: int,
    week_col_start: int,
    week_col_end: int,
    quarter_col: int,
    value_name: str,
    products: list[str]
) -> pd.DataFrame:
    week_names = df.iloc[2, week_col_start:week_col_end].tolist()
    quarter_name = str(df.iloc[1, quarter_col]).strip()

    labels = [f"{str(week).strip()}-{quarter_name}" for week in week_names]
    values = df.iloc[row_start:row_end, week_col_start:week_col_end].values

    records = []
    for i, product in enumerate(products):
        for j, label in enumerate(labels):
            records.append({
                "Product": product,
                "Label": label,
                value_name: values[i][j]
            })

    return pd.DataFrame(records)


def build_weekly_long_table(
    df: pd.DataFrame,
    row_start: int,
    row_end: int,
    value_name: str,
    products: list[str]
) -> pd.DataFrame:
    all_blocks = []

    for block in WEEKLY_BLOCKS:
        block_df = extract_weekly_block(
            df=df,
            row_start=row_start,
            row_end=row_end,
            week_col_start=block["week_start"],
            week_col_end=block["week_end"],
            quarter_col=block["quarter_col"],
            value_name=value_name,
            products=products
        )
        all_blocks.append(block_df)

    return pd.concat(all_blocks, ignore_index=True)


def build_quarterly_long_table(
    df: pd.DataFrame,
    row_start: int,
    row_end: int,
    value_name: str,
    products: list[str]
) -> pd.DataFrame:
    records = []

    for i, product in enumerate(products):
        for q in QUARTERLY_COLS:
            value = df.iloc[row_start + i, q["col"]]
            records.append({
                "Product": product,
                "Label": q["label"],
                value_name: value
            })

    return pd.DataFrame(records)


def build_mpu_weekly_table(df: pd.DataFrame) -> pd.DataFrame:
    records = []

    for product_block in MPU_PRODUCT_BLOCKS:
        product_name = product_block["product"]
        row_start = product_block["row_start"]
        row_end = product_block["row_end"]

        process_names = df.iloc[row_start:row_end, 3].astype(str).str.strip().tolist()

        for block in WEEKLY_BLOCKS:
            week_names = df.iloc[2, block["week_start"]:block["week_end"]].tolist()
            quarter_name = str(df.iloc[1, block["quarter_col"]]).strip()
            labels = [f"{str(week).strip()}-{quarter_name}" for week in week_names]

            values = df.iloc[row_start:row_end, block["week_start"]:block["week_end"]].values

            for i, process_name in enumerate(process_names):
                for j, label in enumerate(labels):
                    records.append({
                        "Product": product_name,
                        "Process": process_name,
                        "Process_Order": i + 1,
                        "Label": label,
                        "MPU": values[i][j],
                        "Report_Type": "MPU_Weekly"
                    })

    mpu_df = pd.DataFrame(records)
    mpu_df["Product"] = mpu_df["Product"].astype(str).str.strip()
    mpu_df["Process"] = mpu_df["Process"].astype(str).str.strip()
    mpu_df["Label"] = mpu_df["Label"].astype(str).str.strip()
    mpu_df["MPU"] = pd.to_numeric(mpu_df["MPU"], errors="coerce")

    split_cols = mpu_df["Label"].str.split("-", n=1, expand=True)
    mpu_df["Week_Label"] = split_cols[0]
    mpu_df["Quarter_Label"] = split_cols[1]
    mpu_df["Week_Num"] = mpu_df["Week_Label"].str.extract(r"(\d+)").astype(float)

    quarter_order = ["WQ326", "WQ426", "WQ127", "WQ227"]
    mpu_df["Quarter_Order"] = mpu_df["Quarter_Label"].apply(
        lambda x: quarter_order.index(x) if x in quarter_order else 999
    )

    mpu_df = mpu_df.sort_values(
        ["Product", "Quarter_Order", "Week_Num", "Process_Order"]
    ).reset_index(drop=True)

    return mpu_df


def enrich_period_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Product"] = df["Product"].astype(str).str.strip()
    df["Label"] = df["Label"].astype(str).str.strip()

    df["Week_Label"] = None
    df["Quarter_Label"] = None
    df["Week_Num"] = None
    df["Quarter_Order"] = None

    weekly_mask = df["Report_Type"] == "Weekly"
    quarterly_mask = df["Report_Type"] == "Quarterly"

    if weekly_mask.any():
        weekly_split = df.loc[weekly_mask, "Label"].str.split("-", n=1, expand=True)
        df.loc[weekly_mask, "Week_Label"] = weekly_split[0]
        df.loc[weekly_mask, "Quarter_Label"] = weekly_split[1]
        df.loc[weekly_mask, "Week_Num"] = df.loc[weekly_mask, "Week_Label"].str.extract(r"(\d+)").astype(float)

    if quarterly_mask.any():
        df.loc[quarterly_mask, "Quarter_Label"] = df.loc[quarterly_mask, "Label"]
        df.loc[quarterly_mask, "Week_Label"] = ""
        df.loc[quarterly_mask, "Week_Num"] = 0

    quarter_order = ["WQ326", "WQ426", "WQ127", "WQ227", "Q326", "Q426", "Q127", "Q227"]
    df["Quarter_Order"] = df["Quarter_Label"].apply(
        lambda x: quarter_order.index(x) if x in quarter_order else 999
    )

    return df


# =============================
# Main process
# =============================
def main() -> None:
    df = pd.read_excel(FILE_NAME, header=None)

    # 1-4 Summary/ Demand
    capacity_weekly_df = build_weekly_long_table(
        df=df,
        row_start=CAPACITY_WEEKLY_ROW_START,
        row_end=CAPACITY_WEEKLY_ROW_END,
        value_name="Capacity",
        products=SUMMARY_PRODUCTS
    )

    demand_weekly_df = build_weekly_long_table(
        df=df,
        row_start=DEMAND_WEEKLY_ROW_START,
        row_end=DEMAND_WEEKLY_ROW_END,
        value_name="Demand",
        products=SUMMARY_PRODUCTS
    )

    weekly_df = pd.merge(
        capacity_weekly_df,
        demand_weekly_df,
        on=["Product", "Label"],
        how="outer"
    )
    weekly_df["Report_Type"] = "Weekly"

    capacity_quarterly_df = build_quarterly_long_table(
        df=df,
        row_start=CAPACITY_QUARTERLY_ROW_START,
        row_end=CAPACITY_QUARTERLY_ROW_END,
        value_name="Capacity",
        products=SUMMARY_PRODUCTS
    )

    demand_quarterly_df = build_quarterly_long_table(
        df=df,
        row_start=DEMAND_QUARTERLY_ROW_START,
        row_end=DEMAND_QUARTERLY_ROW_END,
        value_name="Demand",
        products=SUMMARY_PRODUCTS
    )

    quarterly_df = pd.merge(
        capacity_quarterly_df,
        demand_quarterly_df,
        on=["Product", "Label"],
        how="outer"
    )
    quarterly_df["Report_Type"] = "Quarterly"

    summary_df = pd.concat([weekly_df, quarterly_df], ignore_index=True)
    summary_df["Capacity"] = pd.to_numeric(summary_df["Capacity"], errors="coerce")
    summary_df["Demand"] = pd.to_numeric(summary_df["Demand"], errors="coerce")
    summary_df["Dataset"] = DATASET_NAME
    summary_df = enrich_period_columns(summary_df)
    summary_df = summary_df.sort_values(
        ["Report_Type", "Product", "Quarter_Order", "Week_Num"]
    ).reset_index(drop=True)

    # 6-9 Target
    target_weekly_df = build_weekly_long_table(
        df=df,
        row_start=TARGET_WEEKLY_ROW_START,
        row_end=TARGET_WEEKLY_ROW_END,
        value_name="Capacity_Target",
        products=TARGET_PRODUCTS
    )
    target_weekly_df["Report_Type"] = "Weekly"

    target_quarterly_df = build_quarterly_long_table(
        df=df,
        row_start=TARGET_QUARTERLY_ROW_START,
        row_end=TARGET_QUARTERLY_ROW_END,
        value_name="Capacity_Target",
        products=TARGET_PRODUCTS
    )
    target_quarterly_df["Report_Type"] = "Quarterly"

    target_df = pd.concat([target_weekly_df, target_quarterly_df], ignore_index=True)
    target_df["Capacity_Target"] = pd.to_numeric(target_df["Capacity_Target"], errors="coerce")
    target_df["Dataset"] = DATASET_NAME
    target_df = enrich_period_columns(target_df)
    target_df = target_df.sort_values(
        ["Report_Type", "Product", "Quarter_Order", "Week_Num"]
    ).reset_index(drop=True)

    # 5 MPU
    mpu_df = build_mpu_weekly_table(df)
    mpu_df["Dataset"] = DATASET_NAME

    # Write to Excel with 3 sheets
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="summary", index=False)
        target_df.to_excel(writer, sheet_name="target", index=False)
        mpu_df.to_excel(writer, sheet_name="mpu_weekly", index=False)

    print("✅ Done!")
    print(f"Output file: {OUTPUT_FILE}")
    print("Summary preview:")
    print(summary_df.head(10))
    print("Target preview:")
    print(target_df.head(10))
    print("MPU preview:")
    print(mpu_df.head(10))


if __name__ == "__main__":
    main()