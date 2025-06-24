import streamlit as st
import pandas as pd

st.set_page_config(page_title="紙タオル 使用量比較", layout="wide")
st.title("📋 紙タオル 使用量比較")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df.columns = df.columns.str.strip()

    if "商品コード" not in df.columns:
        df["商品コード"] = "なし"
    else:
        df["商品コード"] = pd.to_numeric(df["商品コード"], errors="coerce")
        df["商品コード"] = df["商品コード"].apply(lambda x: str(int(x)) if pd.notna(x) and x == int(x) else "なし")

    if "原産国" not in df.columns:
        df["原産国"] = ""
    else:
        df["原産国"] = df["原産国"].astype(str).str.strip()

    if "吸水度（分）" not in df.columns:
        df["吸水度（分）"] = ""
    else:
        df["吸水度（分）"] = df["吸水度（分）"].astype(str).str.strip()
        df["吸水度（分）"] = df["吸水度（分）"].replace("nan", "")

    if "略称" not in df.columns:
        st.error("列「略称」が見つかりません。")
        st.stop()

    if "推定使用枚数" in df.columns and "事務所人数" in df.columns:
        df["推定使用枚数"] = pd.to_numeric(df["推定使用枚数"], errors="coerce")
        df["事務所人数"] = pd.to_numeric(df["事務所人数"], errors="coerce")
        df["1日の使用枚数"] = df["推定使用枚数"] / df["事務所人数"]
    else:
        df["1日の使用枚数"] = pd.NA

    df["商品コード_str"] = df["商品コード"]
    df["商品コード_sort"] = df["商品コード_str"].apply(lambda x: int(x) if x != "なし" else float('inf'))

    grouped = df.groupby(["商品コード_str", "略称", "商品コード_sort"], as_index=False).agg({
        "1日の使用枚数": "mean",
        "吸水度（分）": "first",
        "原産国": lambda x: ", ".join(sorted(set(x)))
    })

    grouped["1日の使用枚数"] = grouped["1日の使用枚数"].round(2)
    grouped = grouped[["商品コード_str", "略称", "原産国", "1日の使用枚数", "吸水度（分）", "商品コード_sort"]]
    grouped = grouped.sort_values(by=["商品コード_sort", "略称"], ascending=[True, True]).drop(columns=["商品コード_sort"])
    grouped.rename(columns={"商品コード_str": "商品コード"}, inplace=True)

    selected_items = st.multiselect("略称で絞り込み", options=grouped["略称"].unique(), default=grouped["略称"].unique())
    filtered_df = grouped[grouped["略称"].isin(selected_items)]

    display_df = filtered_df.copy()
    display_df["1日の使用枚数"] = display_df["1日の使用枚数"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "")
    display_df["吸水度（分）"] = display_df["吸水度（分）"].replace("nan", "")
    display_df = display_df.fillna("").astype(str)

    st.dataframe(
        display_df.style.set_table_styles(
            [{'selector': 'th', 'props': [('text-align', 'center')]},
             {'selector': 'td', 'props': [('text-align', 'center')]}]
        ).set_properties(**{'text-align': 'center'}),
        use_container_width=True,
        height=400
    )

    csv = filtered_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ CSVをダウンロード", data=csv, file_name="略称別_使用量一覧.csv", mime="text/csv")
