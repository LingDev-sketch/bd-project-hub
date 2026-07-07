from __future__ import annotations

import io
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
DB_PATH = DATA_DIR / "bd_progress.sqlite"

REGIONS = ["NC1", "NC2", "SW", "SH", "EC", "NE", "NW", "SC", "CC"]
ACTION_TYPES = ["新开", "整改", "关店"]
PROGRESS_OPTIONS = {
    "新开": ["待接洽", "落位沟通中", "商务沟通中", "SLC流程中", "邮件审批中", "已开业", "项目取消"],
    "整改": ["待接洽", "落位洽谈中", "商务洽谈中", "SLC流程中", "邮件审批中", "已开业", "整改取消"],
    "关店": ["沟通中", "邮件审批中", "系统流程中", "已关店", "取消关店"],
}
ADMIN_PASSWORD = "bd-admin"


st.set_page_config(page_title="BD开改关进度管理", page_icon=":material/store:", layout="wide")


def get_conn() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("pragma journal_mode=wal;")
    return conn


def init_db() -> None:
    with get_conn() as conn:
        conn.execute(
            """
            create table if not exists submissions (
                id integer primary key autoincrement,
                project_id text,
                region text,
                action_type text,
                store_code text,
                same_store_code text,
                store_name text,
                plan_month integer,
                plan_date text,
                progress text,
                confirm_status text,
                reason text,
                remark text,
                owner text,
                submitted_by text,
                submitted_at text,
                updated_at text
            )
            """
        )
        conn.execute(
            """
            create table if not exists system_store (
                store_code text,
                same_store_code text,
                store_name text,
                store_status text,
                region text,
                province text,
                city text,
                district text,
                shop_category text,
                sales_group1 text,
                sales_group2 text,
                business_group text,
                first_open_date text,
                actual_open_date text,
                planned_close_date text,
                actual_close_date text,
                refit_start_date text,
                planned_refit_end_date text,
                actual_refit_end_date text,
                area_total real,
                area_sales real,
                imported_at text
            )
            """
        )


def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "0"}:
        return ""
    return text


def normalize_date(value: object) -> str:
    if pd.isna(value) or value == "":
        return ""
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%Y-%m-%d")


def normalize_month(value: object) -> int | None:
    if pd.isna(value) or value == "":
        return None
    try:
        month = int(float(value))
        if 1 <= month <= 12:
            return month
    except Exception:
        return None
    return None


def read_table(table_name: str) -> pd.DataFrame:
    with get_conn() as conn:
        return pd.read_sql_query(f"select * from {table_name}", conn)


def system_lookup() -> pd.DataFrame:
    system = read_table("system_store")
    if system.empty:
        return system
    system["store_code_norm"] = system["store_code"].map(normalize_code)
    system["same_store_code_norm"] = system["same_store_code"].map(normalize_code)
    return system


def import_system_store(file) -> int:
    xl = pd.ExcelFile(file)
    sheet = "校验系统店铺资料" if "校验系统店铺资料" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(file, sheet_name=sheet)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def pick(*cols: str) -> pd.Series:
        for col in cols:
            if col in df.columns:
                return df[col]
        return pd.Series([""] * len(df))

    out = pd.DataFrame(
        {
            "store_code": pick("现用编码", "ERP代码").map(normalize_code),
            "same_store_code": pick("同店编码").map(normalize_code),
            "store_name": pick("店铺名称").fillna("").astype(str).str.strip(),
            "store_status": pick("店铺状态").fillna("").astype(str).str.strip(),
            "region": pick("管理大区(En)", "Region", "Region-1").fillna("").astype(str).str.strip(),
            "province": pick("省份", "Province").fillna("").astype(str).str.strip(),
            "city": pick("地级市", "城市").fillna("").astype(str).str.strip(),
            "district": pick("区县", "市县").fillna("").astype(str).str.strip(),
            "shop_category": pick("店铺大类(En)", "店铺大类").fillna("").astype(str).str.strip(),
            "sales_group1": pick("销售分析客户组1").fillna("").astype(str).str.strip(),
            "sales_group2": pick("销售分析客户组2").fillna("").astype(str).str.strip(),
            "business_group": pick("商业集团").fillna("").astype(str).str.strip(),
            "first_open_date": pick("首次开店日期").map(normalize_date),
            "actual_open_date": pick("实际开店日期").map(normalize_date),
            "planned_close_date": pick("计划关店日期").map(normalize_date),
            "actual_close_date": pick("实际关店日期").map(normalize_date),
            "refit_start_date": pick("整改开始日期").map(normalize_date),
            "planned_refit_end_date": pick("计划整改结束日期").map(normalize_date),
            "actual_refit_end_date": pick("实际整改结束日期").map(normalize_date),
            "area_total": pd.to_numeric(pick("店铺总面积", "总面积"), errors="coerce"),
            "area_sales": pd.to_numeric(pick("营业面积"), errors="coerce"),
            "imported_at": now,
        }
    )
    out = out[(out["store_code"] != "") | (out["same_store_code"] != "") | (out["store_name"] != "")]
    with get_conn() as conn:
        conn.execute("delete from system_store")
        out.to_sql("system_store", conn, if_exists="append", index=False)
    return len(out)


def import_bd_tracking(file) -> int:
    xl = pd.ExcelFile(file)
    rows = []
    mapping = [
        ("提交新开", "新开", "Region", "现用编码", "同店编码", "店铺名称", "BD PLAN month", "开业时间", "进度"),
        ("提交整改", "整改", "Region-1", "现有店编码", "同店编码", "店铺名称", "BD PLAN month", "整改结束时间", "进度"),
        ("提交关店", "关店", "Region-1", "现有店编码", "同店编码", "店铺名称", "BD PLAN month", "关店时间", "进度"),
    ]
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for sheet, action_type, region_col, code_col, same_col, name_col, month_col, date_col, progress_col in mapping:
        if sheet not in xl.sheet_names:
            continue
        df = pd.read_excel(file, sheet_name=sheet)
        for _, row in df.iterrows():
            store_code = normalize_code(row.get(code_col))
            same_code = normalize_code(row.get(same_col))
            store_name = "" if pd.isna(row.get(name_col)) else str(row.get(name_col)).strip()
            if not store_code and not same_code and not store_name:
                continue
            region = "" if pd.isna(row.get(region_col)) else str(row.get(region_col)).strip()
            plan_month = normalize_month(row.get(month_col))
            plan_date = normalize_date(row.get(date_col))
            progress = "" if pd.isna(row.get(progress_col)) else str(row.get(progress_col)).strip()
            project_id = build_project_id(action_type, region, store_code, same_code, store_name)
            rows.append(
                {
                    "project_id": project_id,
                    "region": region,
                    "action_type": action_type,
                    "store_code": store_code,
                    "same_store_code": same_code,
                    "store_name": store_name,
                    "plan_month": plan_month,
                    "plan_date": plan_date,
                    "progress": progress,
                    "confirm_status": "历史导入",
                    "reason": "",
                    "remark": "",
                    "owner": "",
                    "submitted_by": "历史导入",
                    "submitted_at": now,
                    "updated_at": now,
                }
            )
    if not rows:
        return 0
    out = pd.DataFrame(rows)
    with get_conn() as conn:
        conn.execute("delete from submissions where confirm_status = '历史导入'")
        out.to_sql("submissions", conn, if_exists="append", index=False)
    return len(out)


def build_project_id(action_type: str, region: str, store_code: str, same_code: str, store_name: str) -> str:
    key = store_code or same_code or store_name
    return f"{action_type}-{region}-{key}".replace(" ", "")


def save_submission(row: dict[str, object]) -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    store_code = normalize_code(row.get("store_code"))
    same_code = normalize_code(row.get("same_store_code"))
    store_name = str(row.get("store_name") or "").strip()
    row["store_code"] = store_code
    row["same_store_code"] = same_code
    row["store_name"] = store_name
    row["project_id"] = build_project_id(str(row["action_type"]), str(row["region"]), store_code, same_code, store_name)
    row["submitted_at"] = now
    row["updated_at"] = now
    with get_conn() as conn:
        pd.DataFrame([row]).to_sql("submissions", conn, if_exists="append", index=False)


def latest_submissions() -> pd.DataFrame:
    df = read_table("submissions")
    if df.empty:
        return df
    df["submitted_at_dt"] = pd.to_datetime(df["submitted_at"], errors="coerce")
    df = df.sort_values("submitted_at_dt").drop_duplicates("project_id", keep="last")
    return df.drop(columns=["submitted_at_dt"])


def add_system_fields(sub: pd.DataFrame) -> pd.DataFrame:
    system = system_lookup()
    if sub.empty:
        return sub
    out = sub.copy()
    out["store_code_norm"] = out["store_code"].map(normalize_code)
    out["same_store_code_norm"] = out["same_store_code"].map(normalize_code)
    if system.empty:
        out["system_match_status"] = "未导入系统资料"
        return out

    sys_code = system.drop_duplicates("store_code_norm")
    out = out.merge(
        sys_code.add_prefix("sys_"),
        left_on="store_code_norm",
        right_on="sys_store_code_norm",
        how="left",
    )
    unmatched = out["sys_store_name"].isna() & out["same_store_code_norm"].ne("")
    if unmatched.any():
        sys_same = system.drop_duplicates("same_store_code_norm").add_prefix("same_")
        fix = out.loc[unmatched, ["same_store_code_norm"]].merge(
            sys_same,
            left_on="same_store_code_norm",
            right_on="same_same_store_code_norm",
            how="left",
        )
        for col in fix.columns:
            if col.startswith("same_"):
                target = "sys_" + col[5:]
                if target in out.columns:
                    out.loc[unmatched, target] = fix[col].to_numpy()

    out["system_match_status"] = out["sys_store_name"].notna().map({True: "已匹配系统", False: "系统未匹配"})
    return out


def validate_row(row: pd.Series) -> tuple[str, str]:
    action = row.get("action_type", "")
    progress = str(row.get("progress", "") or "")
    store_code = normalize_code(row.get("store_code"))
    store_name = str(row.get("store_name", "") or "").strip()
    plan_date = str(row.get("plan_date", "") or "")
    plan_month = row.get("plan_month")
    sys_status = str(row.get("sys_store_status", "") or "")
    sys_region = str(row.get("sys_region", "") or "")
    sys_open = str(row.get("sys_actual_open_date", "") or "")
    sys_close = str(row.get("sys_actual_close_date", "") or "")
    sys_refit_end = str(row.get("sys_actual_refit_end_date", "") or "")
    sys_matched = row.get("system_match_status") == "已匹配系统"

    issues: list[str] = []
    risk = "通过"

    if not row.get("region"):
        issues.append("缺少Region")
    if not store_name:
        issues.append("缺少店铺名称")
    if not progress:
        issues.append("缺少进度")
    if pd.isna(plan_month) or plan_month in ("", None):
        issues.append("缺少BD PLAN month")
    if not plan_date and progress not in {"项目取消", "整改取消", "取消关店"}:
        issues.append("缺少计划/实际日期")

    if action == "新开":
        if progress in {"已开业"}:
            if not store_code:
                issues.append("新开已开业但无店铺代码")
            if not sys_matched:
                issues.append("新开已开业但系统未匹配")
            if sys_matched and not sys_open:
                issues.append("新开已开业但系统无实际开店日期")
        if progress in {"待接洽", "落位沟通中", "商务沟通中"} and plan_month and int(plan_month) <= 7:
            issues.append("计划月份已到但仍处早期进度，需确认是否可开")
        if progress == "项目取消" and not str(row.get("reason", "") or "").strip():
            issues.append("项目取消但未填写原因")

    if action == "关店":
        if not store_code:
            issues.append("关店必须填写现有店编码")
        if progress == "已关店":
            if not sys_matched:
                issues.append("已关店但系统未匹配")
            if sys_matched and sys_status != "已关店":
                issues.append(f"BD填已关店，但系统状态为{sys_status or '空'}")
            if sys_matched and not sys_close:
                issues.append("BD填已关店，但系统无实际关店日期")
        if progress == "取消关店" and sys_close:
            issues.append("BD填取消关店，但系统已有实际关店日期")

    if action == "整改":
        if not store_code:
            issues.append("整改必须填写现有店编码")
        if progress == "已开业":
            if not sys_matched:
                issues.append("整改完成但系统未匹配")
            if sys_matched and not sys_refit_end:
                issues.append("整改完成但系统无实际整改结束日期")
        if progress == "整改取消" and not str(row.get("reason", "") or "").strip():
            issues.append("整改取消但未填写原因")

    if sys_matched and sys_region and row.get("region") and sys_region != row.get("region"):
        issues.append(f"Region与系统不一致：BD={row.get('region')}，系统={sys_region}")

    if issues:
        severe_words = ["已开业", "已关店", "取消", "系统", "缺少"]
        risk = "高风险" if any(word in "；".join(issues) for word in severe_words) else "需确认"

    return risk, "；".join(issues) if issues else "通过"


def validated_dataset() -> pd.DataFrame:
    sub = latest_submissions()
    if sub.empty:
        return sub
    out = add_system_fields(sub)
    results = out.apply(validate_row, axis=1, result_type="expand")
    out["校验等级"] = results[0]
    out["异常原因"] = results[1]
    useful_cols = [
        "id",
        "project_id",
        "region",
        "action_type",
        "store_code",
        "same_store_code",
        "store_name",
        "plan_month",
        "plan_date",
        "progress",
        "confirm_status",
        "reason",
        "remark",
        "owner",
        "submitted_by",
        "submitted_at",
        "system_match_status",
        "sys_store_status",
        "sys_region",
        "sys_store_name",
        "sys_actual_open_date",
        "sys_actual_close_date",
        "sys_actual_refit_end_date",
        "校验等级",
        "异常原因",
    ]
    return out[[col for col in useful_cols if col in out.columns]]


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    buffer.seek(0)
    return buffer.getvalue()


def submit_page() -> None:
    st.header("BD提报入口")
    st.caption("BD只填关键字段，系统资料能带出的字段不再手工填。")

    with st.form("bd_submit_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            region = st.selectbox("Region", REGIONS)
            action_type = st.selectbox("动作类型", ACTION_TYPES)
            progress = st.selectbox("当前进度", PROGRESS_OPTIONS[action_type])
        with c2:
            store_code = st.text_input("店铺代码/现用编码（没有可空）")
            same_store_code = st.text_input("同店编码（可空）")
            store_name = st.text_input("店铺名称")
        with c3:
            plan_month = st.selectbox("BD PLAN month", list(range(1, 13)))
            plan_date = st.date_input("计划/实际日期")
            owner = st.text_input("负责人")

        confirm_status = st.radio("是否确认继续推进", ["确认推进", "暂缓/待确认", "取消/不推进"], horizontal=True)
        reason = st.text_area("取消/延迟/异常原因")
        remark = st.text_area("最新备注")
        submitted_by = st.text_input("提交人")
        submitted = st.form_submit_button("提交")

    if submitted:
        save_submission(
            {
                "project_id": "",
                "region": region,
                "action_type": action_type,
                "store_code": store_code,
                "same_store_code": same_store_code,
                "store_name": store_name,
                "plan_month": plan_month,
                "plan_date": plan_date.strftime("%Y-%m-%d"),
                "progress": progress,
                "confirm_status": confirm_status,
                "reason": reason,
                "remark": remark,
                "owner": owner,
                "submitted_by": submitted_by,
                "submitted_at": "",
                "updated_at": "",
            }
        )
        st.success("已提交。后台会自动匹配系统资料并生成校验结果。")


def admin_page() -> None:
    st.header("后台管理")
    password = st.text_input("后台密码", type="password")
    if password != ADMIN_PASSWORD:
        st.info("请输入后台密码。默认密码在 app.py 顶部，可自行修改。")
        return

    st.success("已进入后台。")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. 导入系统店铺资料")
        sys_file = st.file_uploader("上传包含“校验系统店铺资料”的Excel", type=["xlsx"], key="sys_file")
        if sys_file and st.button("导入系统资料"):
            count = import_system_store(sys_file)
            st.success(f"已导入系统店铺资料 {count:,} 行。")
    with c2:
        st.subheader("2. 可选：导入历史BD Tracking")
        bd_file = st.file_uploader("上传历史BD Tracking Excel", type=["xlsx"], key="bd_file")
        if bd_file and st.button("导入历史三张提交表"):
            count = import_bd_tracking(bd_file)
            st.success(f"已导入历史提交记录 {count:,} 行。")

    data = validated_dataset()
    if data.empty:
        st.warning("暂无BD提报数据。")
        return

    st.divider()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("项目数", f"{len(data):,}")
    k2.metric("高风险", f"{(data['校验等级'] == '高风险').sum():,}")
    k3.metric("需确认", f"{(data['校验等级'] == '需确认').sum():,}")
    k4.metric("已通过", f"{(data['校验等级'] == '通过').sum():,}")

    with st.expander("筛选", expanded=True):
        fc1, fc2, fc3 = st.columns(3)
        selected_regions = fc1.multiselect("Region", REGIONS, default=REGIONS)
        selected_actions = fc2.multiselect("动作类型", ACTION_TYPES, default=ACTION_TYPES)
        selected_levels = fc3.multiselect("校验等级", ["高风险", "需确认", "通过"], default=["高风险", "需确认", "通过"])

    view = data[
        data["region"].isin(selected_regions)
        & data["action_type"].isin(selected_actions)
        & data["校验等级"].isin(selected_levels)
    ].copy()

    st.subheader("项目明细与校验结果")
    st.dataframe(view, use_container_width=True, hide_index=True)

    st.subheader("区域进度汇总")
    summary = (
        data.pivot_table(
            index=["region", "action_type"],
            columns="progress",
            values="project_id",
            aggfunc="count",
            fill_value=0,
        )
        .reset_index()
    )
    st.dataframe(summary, use_container_width=True, hide_index=True)

    exceptions = data[data["校验等级"].isin(["高风险", "需确认"])].copy()
    export_bytes = to_excel_bytes(
        {
            "全国项目库": data,
            "异常清单": exceptions,
            "区域进度汇总": summary,
        }
    )
    st.download_button(
        "下载全国汇总与异常清单Excel",
        export_bytes,
        file_name=f"BD开改关进度校验_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def rules_page() -> None:
    st.header("自动校验规则")
    st.markdown(
        """
**可信优先级**

系统店铺资料 > BD提报资料 > BD备注解释

**新开**

- 待接洽、落位沟通中、商务沟通中：允许没有店铺代码，但必须有店铺名称、计划月份、进度和备注。
- 已开业：必须有店铺代码，且应能匹配系统资料，系统应有实际开店日期。
- 计划月份已到但仍处于早期进度，会标记为需确认。
- 项目取消必须填写取消原因。

**关店**

- 关店必须有现有店编码。
- 已关店必须能匹配系统资料，且系统状态应为已关店，并有实际关店日期。
- 取消关店但系统已有实际关店日期，会标记为高风险。

**整改**

- 整改必须有现有店编码。
- 已开业/整改完成应匹配系统资料，并有实际整改结束日期。
- 整改取消必须填写原因。

**跨表一致性**

- BD填报Region与系统管理大区不一致，会标记异常。
- 系统未匹配的记录会进入异常清单，便于追问BD。
        """
    )


def main() -> None:
    init_db()
    st.title("BD开改关进度管理与自动校验系统")
    st.caption("让BD只填关键进度，让系统自动匹配店铺资料、发现异常、导出全国汇总。")
    page = st.sidebar.radio("页面", ["BD提报入口", "后台管理", "校验规则说明"])
    if page == "BD提报入口":
        submit_page()
    elif page == "后台管理":
        admin_page()
    else:
        rules_page()


if __name__ == "__main__":
    main()
