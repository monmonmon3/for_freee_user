import pandas as pd
from typing import List, Union

def insert_summary_rows(
    df: pd.DataFrame,
    selected_departments: Union[List[str], str]
) -> pd.DataFrame:
    dept = selected_departments[0] if isinstance(selected_departments, list) else selected_departments

    if not (df["小分類"] == "売上高").any():
        new_row = pd.Series(0, index=df.columns)
        new_row["勘定科目"] = "純売上高"
        new_row["小分類"]   = "純売上高"
        new_row["部門"]     = dept
        df = pd.concat([pd.DataFrame([new_row]), df], ignore_index=True)

    if not (df["小分類"] == "当期商品仕入").any():
        idxs = df.index[df["小分類"] == "純売上高"]
        if idxs.empty:
            idxs = df.index[df["小分類"] == "売上高"]
        if not idxs.empty:
            insert_at = idxs[0]
            row_begin = pd.Series(0, index=df.columns)
            row_begin["勘定科目"] = "期首商品棚卸"
            row_begin["小分類"]   = "期首商品棚卸"
            row_begin["部門"]     = dept
            row_pur = pd.Series(0, index=df.columns)
            row_pur["勘定科目"] = "純仕入高"
            row_pur["小分類"]   = "純仕入高"
            row_pur["部門"]     = dept

            top    = df.iloc[:insert_at+1]
            bottom = df.iloc[insert_at+1:]
            df = pd.concat([
                top,
                pd.DataFrame([row_begin, row_pur]),
                bottom
            ], ignore_index=True)

    def _insert_summary_row(tmp: pd.DataFrame, targets: List[str], label: str) -> pd.DataFrame:
        subset = tmp[tmp["小分類"].isin(targets)]
        if subset.empty:
            return tmp
        last_idx = subset.index[-1]
        sums = subset.select_dtypes(include="number").sum()
        row = pd.Series(sums.to_dict(), name="summary")
        row["勘定科目"] = label
        row["小分類"]   = label
        row["部門"]     = dept
        top    = tmp.loc[:last_idx]
        bottom = tmp.loc[last_idx+1:]
        return pd.concat([top, pd.DataFrame([row]), bottom], ignore_index=True)

    for targets, label in [
        (["売上高"],       "純売上高"),
        (["当期商品仕入"], "純仕入高"),
        (["人件費"],       "人件費合計"),
        (["人件費", "販売管理費"], "販売費及び一般管理費"),
        (["営業外収益"],   "営業外収益合計"),
        (["営業外費用"],   "営業外費用合計"),
        (["特別利益"],     "特別利益合計"),
        (["特別損失"],     "特別損失合計"),
        (["法人税等"],     "法人税・住民税・事業税"),
    ]:
        df = _insert_summary_row(df, targets, label)

    def _ensure_row(tmp: pd.DataFrame, prev_label: str, new_label: str) -> pd.DataFrame:
        values = tmp["勘定科目"].values
        if prev_label in values and new_label not in values:
            idxs = tmp.index[tmp["勘定科目"] == prev_label]
            if idxs.any():
                last_idx = idxs[-1]
                new = pd.Series(0, index=tmp.columns)
                new["勘定科目"] = new_label
                new["小分類"]   = new_label
                new["部門"]     = dept
                top    = tmp.iloc[:last_idx+1]
                bottom = tmp.iloc[last_idx+1:]
                return pd.concat([top, pd.DataFrame([new]), bottom], ignore_index=True)
        return tmp

    for prev, new in [
        ("純仕入高",           "期末商品棚卸"),
        ("純売上高",           "期首商品棚卸"),
        ("期末商品棚卸",       "売上原価"),
        ("売上原価",           "売上総利益"),
        ("販売費及び一般管理費", "営業利益"),
        ("営業利益",           "営業外収益合計"),
        ("営業外収益合計",     "営業外費用合計"),
        ("営業外費用合計",     "経常利益"),
        ("経常利益",           "特別利益合計"),
        ("特別利益合計",       "特別損失合計"),
        ("特別損失合計",       "税引前当期純利益"),
        ("税引前当期純利益",   "法人税・住民税・事業税"),
        ("法人税・住民税・事業税", "税引後当期純利益"),
    ]:
        df = _ensure_row(df, prev, new)

    return df


def calculate_financials(df: pd.DataFrame) -> pd.DataFrame:
    """
    各月の売上原価、売上総利益、営業利益、経常利益、
    税引前当期純利益、税引後当期純利益を計算して反映します。
    """
    cols = df.columns[4:]  # 月次数値列

    # 売上原価 = 期首商品棚卸 + 純仕入高 - 期末商品棚卸
    if all(x in df["勘定科目"].values for x in ["期首商品棚卸", "純仕入高", "売上原価"]):
        i0 = df.index[df["勘定科目"] == "期首商品棚卸"][0]
        i1 = df.index[df["勘定科目"] == "純仕入高"][0]
        i2 = df.index[df["勘定科目"] == "売上原価"][0]
        for c in cols:
            b = pd.to_numeric(df.at[i0, c], errors="coerce")
            p = pd.to_numeric(df.at[i1, c], errors="coerce")
            e = pd.to_numeric(df.at[i2, c], errors="coerce")
            df.at[i2, c] = b + p - e

    # 売上総利益 = 純売上高 - 売上原価
    if all(x in df["勘定科目"].values for x in ["純売上高", "売上原価", "売上総利益"]):
        i0 = df.index[df["勘定科目"] == "純売上高"][0]
        i1 = df.index[df["勘定科目"] == "売上原価"][0]
        i2 = df.index[df["勘定科目"] == "売上総利益"][0]
        for c in cols:
            s = pd.to_numeric(df.at[i0, c], errors="coerce")
            k = pd.to_numeric(df.at[i1, c], errors="coerce")
            df.at[i2, c] = s - k

    # 営業利益 = 売上総利益 - 販管費
    if all(x in df["勘定科目"].values for x in ["売上総利益", "販売費及び一般管理費", "営業利益"]):
        i0 = df.index[df["勘定科目"] == "売上総利益"][0]
        i1 = df.index[df["勘定科目"] == "販売費及び一般管理費"][0]
        i2 = df.index[df["勘定科目"] == "営業利益"][0]
        for c in cols:
            g = pd.to_numeric(df.at[i0, c], errors="coerce")
            h = pd.to_numeric(df.at[i1, c], errors="coerce")
            df.at[i2, c] = g - h

    # 経常利益 = 営業利益 + 営業外収益 - 営業外費用
    if all(x in df["勘定科目"].values for x in ["営業利益", "営業外収益合計", "営業外費用合計", "経常利益"]):
        i0 = df.index[df["勘定科目"] == "営業利益"][0]
        i1 = df.index[df["勘定科目"] == "営業外収益合計"][0]
        i2 = df.index[df["勘定科目"] == "営業外費用合計"][0]
        i3 = df.index[df["勘定科目"] == "経常利益"][0]
        for c in cols:
            o0 = pd.to_numeric(df.at[i0, c], errors="coerce")
            o1 = pd.to_numeric(df.at[i1, c], errors="coerce")
            o2 = pd.to_numeric(df.at[i2, c], errors="coerce")
            df.at[i3, c] = o0 + o1 - o2

    # 税引前利益 = 経常利益 + 特別利益 - 特別損失
    if all(x in df["勘定科目"].values for x in ["経常利益", "特別利益合計", "特別損失合計", "税引前当期純利益"]):
        i0 = df.index[df["勘定科目"] == "経常利益"][0]
        i1 = df.index[df["勘定科目"] == "特別利益合計"][0]
        i2 = df.index[df["勘定科目"] == "特別損失合計"][0]
        i3 = df.index[df["勘定科目"] == "税引前当期純利益"][0]
        for c in cols:
            v0 = pd.to_numeric(df.at[i0, c], errors="coerce")
            v1 = pd.to_numeric(df.at[i1, c], errors="coerce")
            v2 = pd.to_numeric(df.at[i2, c], errors="coerce")
            df.at[i3, c] = v0 + v1 - v2

    # 税引後利益 = 税引前利益 - 税金
    if all(x in df["勘定科目"].values for x in ["税引前当期純利益", "法人税・住民税・事業税", "税引後当期純利益"]):
        i0 = df.index[df["勘定科目"] == "税引前当期純利益"][0]
        i1 = df.index[df["勘定科目"] == "法人税・住民税・事業税"][0]
        i2 = df.index[df["勘定科目"] == "税引後当期純利益"][0]
        for c in cols:
            p = pd.to_numeric(df.at[i0, c], errors="coerce")
            t = pd.to_numeric(df.at[i1, c], errors="coerce")
            df.at[i2, c] = p - t

    return df
