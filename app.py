import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Capacity Dashboard", layout="wide")

# =============================
# Load data
# =============================
summary_df = pd.read_excel("output/capacity_dashboard_dataset.xlsx", sheet_name="summary")
target_df = pd.read_excel("output/capacity_dashboard_dataset.xlsx", sheet_name="target")
mpu_df = pd.read_excel("output/capacity_dashboard_dataset.xlsx", sheet_name="mpu_weekly")

# =============================
# Clean data
# =============================
for df in [summary_df, target_df, mpu_df]:
    df["Product"] = df["Product"].astype(str).str.strip()
    if "Label" in df.columns:
        df["Label"] = df["Label"].astype(str).str.strip()

summary_df["Capacity"] = pd.to_numeric(summary_df["Capacity"], errors="coerce")
summary_df["Demand"] = pd.to_numeric(summary_df["Demand"], errors="coerce")
summary_df["Report_Type"] = summary_df["Report_Type"].astype(str).str.strip()
summary_df["Quarter_Label"] = summary_df["Quarter_Label"].astype(str).str.strip()

target_df["Capacity_Target"] = pd.to_numeric(target_df["Capacity_Target"], errors="coerce")
target_df["Report_Type"] = target_df["Report_Type"].astype(str).str.strip()
target_df["Quarter_Label"] = target_df["Quarter_Label"].astype(str).str.strip()

mpu_df["Process"] = mpu_df["Process"].astype(str).str.strip()
mpu_df["MPU"] = pd.to_numeric(mpu_df["MPU"], errors="coerce")
mpu_df["Process_Order"] = pd.to_numeric(mpu_df["Process_Order"], errors="coerce")
mpu_df["Quarter_Label"] = mpu_df["Quarter_Label"].astype(str).str.strip()

# =============================
# Dataset field
# =============================
if "Dataset" not in summary_df.columns:
    summary_df["Dataset"] = "HPL_HRS"

if "Dataset" not in target_df.columns:
    target_df["Dataset"] = "HPL_HRS"

if "Dataset" not in mpu_df.columns:
    mpu_df["Dataset"] = "HPL_HRS"

summary_df["Dataset"] = summary_df["Dataset"].astype(str).str.strip()
target_df["Dataset"] = target_df["Dataset"].astype(str).str.strip()
mpu_df["Dataset"] = mpu_df["Dataset"].astype(str).str.strip()
# =============================
# UI
# =============================
st.title("📊 Capacity Dashboard")

dataset_list = sorted(
    set(summary_df["Dataset"].dropna().unique().tolist())
    | set(target_df["Dataset"].dropna().unique().tolist())
    | set(mpu_df["Dataset"].dropna().unique().tolist())
)

default_dataset_index = dataset_list.index("HPL_HRS") if "HPL_HRS" in dataset_list else 0

selected_dataset = st.selectbox(
    "Select Dataset",
    dataset_list,
    index=default_dataset_index
)

analysis_type = st.selectbox(
    "Analysis Type",
    [
        "Capacity Summary",
        "Capacity Summary and Demand",
        "Capacity Target",
        "Capacity Summary and Capacity Target",
        "Cycle Time (Minute per Unit)"
    ]
)

period = None
if analysis_type != "Cycle Time (Minute per Unit)":
    period = st.selectbox("Period", ["Weekly", "Quarterly"])

if analysis_type == "Capacity Summary":
    report_type = f"Capacity Summary by {period}"
elif analysis_type == "Capacity Summary and Demand":
    report_type = f"Capacity Summary and Demand by {period}"
elif analysis_type == "Capacity Target":
    report_type = f"Capacity Target by {period}"
elif analysis_type == "Capacity Summary and Capacity Target":
    report_type = f"Capacity Summary and Capacity Target by {period}"
else:
    report_type = "Cycle Time (Minute per Unit, MPU) by Process by Weekly"

# =============================
# Colors
# =============================
COLORS = {
    "WQ326": "#7BC96F",
    "WQ426": "#FFB6C1",
    "WQ127": "#6BAED6",
    "WQ227": "#FDBE85",
    "Q326": "#7BC96F",
    "Q426": "#FFB6C1",
    "Q127": "#6BAED6",
    "Q227": "#FDBE85",
}

def get_colors(df: pd.DataFrame) -> list[str]:
    return [COLORS.get(x, "#999999") for x in df["Quarter_Label"]]

def sort_by_label_list(df: pd.DataFrame, label_sort: list[str]) -> pd.DataFrame:
    order_map = {label: i for i, label in enumerate(label_sort)}
    out = df.copy()
    out["_sort_order"] = out["Label"].map(order_map)
    out = out.sort_values("_sort_order").drop(columns=["_sort_order"]).reset_index(drop=True)
    out["Label"] = out["Label"].astype(str)
    return out

def style_fig(fig: go.Figure, y_title: str, x_title: str = "Period") -> go.Figure:
    fig.update_layout(
        height=500,
        plot_bgcolor="white",
        paper_bgcolor="white",
        hovermode="closest",
        margin=dict(l=20, r=10, t=20, b=20),
        bargap=0.10,
        barcornerradius=18,
        xaxis=dict(
            title=x_title,
            tickangle=-45,
            showgrid=False,
            automargin=True
        ),
        yaxis=dict(
            title=y_title,
            showgrid=True,
            gridcolor="#E5E7EB",
            zeroline=False,
            automargin=True
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
    )
    return fig

def build_bar_chart(df: pd.DataFrame, y_col: str, y_title: str, label_sort: list[str]) -> go.Figure:
    chart_df = sort_by_label_list(df, label_sort)
    x_pos = list(range(len(chart_df)))

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=x_pos,
            y=chart_df[y_col].tolist(),
            marker=dict(
                color=get_colors(chart_df),
                line=dict(width=0)
            ),
            opacity=0.65,
            name=y_title,
            width=0.74,
            hovertemplate=(
                "Label: %{customdata[0]}<br>"
                f"{y_title}: %{{y:,.2f}}"
                "<extra></extra>"
            ),
            customdata=chart_df[["Label"]].values
        )
    )

    fig = style_fig(fig, y_title=y_title)
    fig.update_xaxes(
        tickmode="array",
        tickvals=x_pos,
        ticktext=chart_df["Label"].tolist(),
        range=[-0.5, len(chart_df) - 0.5]
    )
    return fig

def build_comparison_chart(
    df: pd.DataFrame,
    line_col: str,
    line_name: str,
    line_color: str,
    alert_mask: pd.Series,
    label_sort: list[str]
) -> go.Figure:
    chart_df = sort_by_label_list(df, label_sort)
    alert_mask = alert_mask.loc[chart_df.index].reset_index(drop=True)
    x_pos = list(range(len(chart_df)))

    chart_df["Gap_Signed"] = chart_df["Capacity"] - chart_df[line_col]
    chart_df["Gap_Label"] = chart_df["Gap_Signed"].map(
        lambda x: f"{x:+,.2f}" if pd.notna(x) else ""
    )

    fig = go.Figure()

    # Bar = Capacity Summary
    fig.add_trace(
        go.Bar(
            x=x_pos,
            y=chart_df["Capacity"].tolist(),
            marker=dict(
                color=get_colors(chart_df),
                line=dict(width=0)
            ),
            opacity=0.65,
            name="Capacity Summary",
            width=0.74,
            hovertemplate=(
                "Label: %{customdata[0]}<br>"
                "Capacity Summary: %{y:,.2f}<br>"
                "Gap (+/-): %{customdata[1]}"
                "<extra></extra>"
            ),
            customdata=chart_df[["Label", "Gap_Label"]].values
        )
    )

    # Line
    fig.add_trace(
        go.Scatter(
            x=x_pos,
            y=chart_df[line_col].tolist(),
            mode="lines",
            line=dict(color=line_color, width=3),
            name=line_name,
            hovertemplate=(
                "Label: %{customdata[0]}<br>"
                f"{line_name}: %{{y:,.2f}}<br>"
                "Gap (+/-): %{customdata[1]}"
                "<extra></extra>"
            ),
            customdata=chart_df[["Label", "Gap_Label"]].values
        )
    )

    # Normal markers
    normal_df = chart_df[~alert_mask].copy().reset_index(drop=True)
    normal_x = [x_pos[i] for i, flag in enumerate(~alert_mask) if flag]

    fig.add_trace(
        go.Scatter(
            x=normal_x,
            y=normal_df[line_col].tolist(),
            mode="markers",
            marker=dict(
                size=7,
                color="white",
                line=dict(color=line_color, width=2)
            ),
            name=f"{line_name} OK",
            hovertemplate=(
                "Label: %{customdata[0]}<br>"
                f"{line_name}: %{{y:,.2f}}<br>"
                "Gap (+/-): %{customdata[1]}"
                "<extra></extra>"
            ),
            customdata=normal_df[["Label", "Gap_Label"]].values,
            showlegend=False
        )
    )

    # Alert markers
    alert_df = chart_df[alert_mask].copy().reset_index(drop=True)
    alert_x = [x_pos[i] for i, flag in enumerate(alert_mask) if flag]

    fig.add_trace(
        go.Scatter(
            x=alert_x,
            y=alert_df[line_col].tolist(),
            mode="markers",
            marker=dict(
                size=8,
                color="white",
                line=dict(color="red", width=2.5)
            ),
            name=f"{line_name} Alert",
            hovertemplate=(
                "Label: %{customdata[0]}<br>"
                f"{line_name}: %{{y:,.2f}}<br>"
                "Gap (+/-): %{customdata[1]}"
                "<extra></extra>"
            ),
            customdata=alert_df[["Label", "Gap_Label"]].values,
            showlegend=False
        )
    )

    y_max = max(chart_df["Capacity"].max(), chart_df[line_col].max()) * 1.12

    fig = style_fig(fig, y_title="Value")
    fig.update_xaxes(
        tickmode="array",
        tickvals=x_pos,
        ticktext=chart_df["Label"].tolist(),
        range=[-0.5, len(chart_df) - 0.5]
    )
    fig.update_yaxes(range=[0, y_max])

    return fig

# =============================
# Report: MPU
# =============================
if report_type == "Cycle Time (Minute per Unit, MPU) by Process by Weekly":
    current_mpu_df = mpu_df[mpu_df["Dataset"] == selected_dataset].copy()
    product_list = sorted(current_mpu_df["Product"].dropna().unique().tolist())
    default_index = product_list.index("Hoplite (FBG) 980") if "Hoplite (FBG) 980" in product_list else 0

    selected_product = st.selectbox("Select Product", product_list, index=default_index)

    product_df = current_mpu_df[current_mpu_df["Product"] == selected_product].copy()
    week_list = product_df["Label"].drop_duplicates().tolist()
    selected_week = st.selectbox("Select Week", week_list, index=0 if week_list else None)

    filtered_mpu_df = product_df[product_df["Label"] == selected_week].copy().sort_values("Process_Order")

    if not filtered_mpu_df.empty:
        avg_mpu = filtered_mpu_df["MPU"].mean()
        max_mpu = filtered_mpu_df["MPU"].max()
        min_mpu = filtered_mpu_df["MPU"].min()

        col1, col2, col3 = st.columns(3)
        col1.metric("Average MPU", f"{avg_mpu:,.2f}")
        col2.metric("Max MPU", f"{max_mpu:,.2f}")
        col3.metric("Min MPU", f"{min_mpu:,.2f}")

        st.subheader(f"{report_type} - {selected_product} - {selected_week}")

        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=filtered_mpu_df["Process"].tolist(),
                y=filtered_mpu_df["MPU"].tolist(),
                marker=dict(
                    color="#808080",
                    line=dict(width=0)
                ),
                opacity=0.65,
                width=0.68,
                hovertemplate="Process: %{x}<br>MPU: %{y:,.2f}<extra></extra>"
            )
        )
        fig = style_fig(fig, y_title="MPU", x_title="Process")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Detail Data")
        display_df = filtered_mpu_df[["Product", "Process", "Label", "MPU"]].copy()
        display_df["MPU"] = display_df["MPU"].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
        st.dataframe(display_df, width="stretch")
    else:
        st.warning("No MPU data found.")

# =============================
# Reports: Summary / Summary+Demand
# =============================
elif report_type in [
    "Capacity Summary by Weekly",
    "Capacity Summary by Quarterly",
    "Capacity Summary and Demand by Weekly",
    "Capacity Summary and Demand by Quarterly"
]:
    current_period = "Weekly" if "Weekly" in report_type else "Quarterly"
    current_df = summary_df[
        (summary_df["Report_Type"] == current_period) &
        (summary_df["Dataset"] == selected_dataset)
    ].copy()

    product_list = sorted(current_df["Product"].dropna().unique().tolist())
    default_index = product_list.index("Total Capacity") if "Total Capacity" in product_list else 0
    selected_product = st.selectbox("Select Product", product_list, index=default_index)

    filtered_df = current_df[current_df["Product"] == selected_product].copy()

    if not filtered_df.empty:
        label_sort = filtered_df["Label"].drop_duplicates().tolist()
        filtered_df = sort_by_label_list(filtered_df, label_sort)

        total_capacity = filtered_df["Capacity"].sum()

        if "Demand" in report_type:
            col1, col2 = st.columns(2)
            col1.metric("Selected Product", selected_product)
            col2.metric("Periods", f"{len(label_sort):,}")
        else:
            col1, col2 = st.columns(2)
            col1.metric("Total Capacity", f"{total_capacity:,.2f}")
            col2.metric("Selected Product", selected_product)

        st.subheader(f"{report_type} - {selected_product}")

        if "Demand" not in report_type:
            fig = build_bar_chart(filtered_df, y_col="Capacity", y_title="Capacity Summary", label_sort=label_sort)
        else:
            st.caption("Legend: bar chart = Capacity Summary | line chart = Demand")
            alert_mask = filtered_df["Capacity"] < filtered_df["Demand"]
            fig = build_comparison_chart(
                filtered_df,
                line_col="Demand",
                line_name="Demand",
                line_color="black",
                alert_mask=alert_mask,
                label_sort=label_sort
            )

        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Detail Data")
        if "Demand" in report_type:
            display_df = filtered_df[["Product", "Label", "Capacity", "Demand", "Quarter_Label"]].copy()
            display_df["Gap (+/-)"] = filtered_df["Capacity"] - filtered_df["Demand"]
            display_df["Capacity"] = display_df["Capacity"].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
            display_df["Demand"] = display_df["Demand"].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
            display_df["Gap (+/-)"] = display_df["Gap (+/-)"].map(
                lambda x: f"{x:+,.2f}" if pd.notna(x) else ""
            )
            display_df = display_df[["Product", "Label", "Capacity", "Demand", "Gap (+/-)", "Quarter_Label"]]
        else:
            display_df = filtered_df[["Product", "Label", "Capacity", "Quarter_Label"]].copy()
            display_df["Capacity"] = display_df["Capacity"].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")

        st.dataframe(display_df, width="stretch")
    else:
        st.warning("No data found.")

# =============================
# Reports: Target / Summary+Target
# =============================
else:
    current_period = "Weekly" if "Weekly" in report_type else "Quarterly"
    current_target_df = target_df[
        (target_df["Report_Type"] == current_period) &
        (target_df["Dataset"] == selected_dataset)
    ].copy()

    product_list = sorted(current_target_df["Product"].dropna().unique().tolist())
    default_index = product_list.index("Total build plan") if "Total build plan" in product_list else 0
    selected_product = st.selectbox("Select Product", product_list, index=default_index)

    filtered_target_df = current_target_df[current_target_df["Product"] == selected_product].copy()

    if not filtered_target_df.empty:
        label_sort = filtered_target_df["Label"].drop_duplicates().tolist()
        filtered_target_df = sort_by_label_list(filtered_target_df, label_sort)
        total_target = filtered_target_df["Capacity_Target"].sum()

        if report_type in [
            "Capacity Summary and Capacity Target by Weekly",
            "Capacity Summary and Capacity Target by Quarterly"
        ]:
            summary_period_df = summary_df[
                (summary_df["Report_Type"] == current_period) &
                (summary_df["Dataset"] == selected_dataset)
            ].copy()
            summary_product = "Total Capacity" if selected_product == "Total build plan" else selected_product
            filtered_summary_df = summary_period_df[summary_period_df["Product"] == summary_product].copy()

            merged_df = pd.merge(
                filtered_target_df[["Product", "Label", "Quarter_Label", "Capacity_Target"]],
                filtered_summary_df[["Product", "Label", "Capacity"]],
                on="Label",
                how="left",
                suffixes=("_target", "_summary")
            )

            merged_df = sort_by_label_list(merged_df, label_sort)
            alert_mask = merged_df["Capacity"] < merged_df["Capacity_Target"]

            col1, col2 = st.columns(2)
            col1.metric("Selected Product", selected_product)
            col2.metric("Periods", f"{len(label_sort):,}")

            st.subheader(f"{report_type} - {selected_product}")
            st.caption("Legend: bar chart = Capacity Summary | line chart = Capacity Target")

            fig = build_comparison_chart(
                merged_df,
                line_col="Capacity_Target",
                line_name="Capacity Target",
                line_color="green",
                alert_mask=alert_mask,
                label_sort=label_sort
            )

            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Detail Data")
            display_df = merged_df[[
                "Product_target", "Label", "Quarter_Label", "Capacity_Target", "Product_summary", "Capacity"
            ]].copy()
            display_df["Gap (+/-)"] = merged_df["Capacity"] - merged_df["Capacity_Target"]
            display_df["Capacity"] = display_df["Capacity"].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
            display_df["Capacity_Target"] = display_df["Capacity_Target"].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
            display_df["Gap (+/-)"] = display_df["Gap (+/-)"].map(
                lambda x: f"{x:+,.2f}" if pd.notna(x) else ""
            )
            st.dataframe(display_df, width="stretch")

        else:
            col1, col2 = st.columns(2)
            col1.metric("Total Capacity Target", f"{total_target:,.2f}")
            col2.metric("Selected Product", selected_product)

            st.subheader(f"{report_type} - {selected_product}")

            fig = build_bar_chart(filtered_target_df, y_col="Capacity_Target", y_title="Capacity Target", label_sort=label_sort)
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Detail Data")
            display_df = filtered_target_df[["Product", "Label", "Quarter_Label", "Capacity_Target"]].copy()
            display_df["Capacity_Target"] = display_df["Capacity_Target"].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
            st.dataframe(display_df, width="stretch")
    else:
        st.warning("No data found.")