import streamlit as st
import pandas as pd
from supabase import create_client, Client
from datetime import datetime, date
from io import BytesIO
import openpyxl
import os
import calendar
import altair as alt

st.set_page_config(page_title="Insurance Commission Tracker", layout="wide")

SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_ANON_KEY"]
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

menu = st.sidebar.radio("Navigate", [
    "1. Daily Sales Entry and Current Month Sales",
    "2. Search and Edit Sales with Date Range",
    "3. Upload Commission, Detect Duplicates, Upload History",
    "4. Yearly Sales Chart"
])

st.title("Insurance Commission Tracker")

response = supabase.table("daily_sales").select("*").execute()
df_sales = pd.DataFrame(response.data)

today = pd.to_datetime("today").normalize()

if not df_sales.empty:
    df_sales["date_of_sale"] = pd.to_datetime(df_sales["date_of_sale"], errors="coerce")
    df_sales["effective_date"] = pd.to_datetime(df_sales["effective_date"], errors="coerce")
    df_sales["premium"] = pd.to_numeric(df_sales["premium"], errors="coerce")
    df_sales["month"] = df_sales["date_of_sale"].dt.month
    df_sales["year"] = df_sales["date_of_sale"].dt.year

if menu == "1. Daily Sales Entry and Current Month Sales":
    current_month_sales = df_sales[
        (df_sales["date_of_sale"].dt.month == today.month) &
        (df_sales["date_of_sale"].dt.year == today.year)
    ].copy()

    total_agency = current_month_sales["premium"].sum()

    chart_col, summary_col = st.columns([3, 1])

    with chart_col:
        if not current_month_sales.empty:
            agent_totals = (
                current_month_sales.groupby("agent_name", dropna=False)["premium"]
                .sum().reset_index().rename(columns={"premium": "sum"})
            )

            month_start = pd.Timestamp(today.year, today.month, 1)
            month_end = pd.Timestamp(today.year, today.month, calendar.monthrange(today.year, today.month)[1])
            bdays_elapsed = len(pd.bdate_range(month_start, today))
            bdays_total = len(pd.bdate_range(month_start, month_end))
            bdays_elapsed = max(bdays_elapsed, 1)

            weekday_to_date = current_month_sales[
                (current_month_sales["date_of_sale"] <= today) &
                (current_month_sales["date_of_sale"].dt.dayofweek < 5)
            ].copy()

            pace_totals = (
                weekday_to_date.groupby("agent_name", dropna=False)["premium"]
                .sum().reset_index().rename(columns={"premium": "pace_sum"})
            )

            merged = agent_totals.merge(pace_totals, on="agent_name", how="left")
            merged["pace_sum"] = merged["pace_sum"].fillna(0)
            merged["forecast"] = (merged["pace_sum"] / bdays_elapsed) * bdays_total

            plot_df = pd.melt(
                merged,
                id_vars=["agent_name"],
                value_vars=["sum", "forecast"],
                var_name="type",
                value_name="value"
            ).replace({"type": {"sum": "Current", "forecast": "Forecast"}})

            ch = (
                alt.Chart(plot_df)
                .mark_bar()
                .encode(
                    x=alt.X("agent_name:N", title=None,
                            axis=alt.Axis(labelAngle=45, labelFontSize=9, titleFontSize=10)),
                    y=alt.Y("value:Q", title="Premium"),
                    color=alt.Color("type:N", scale=alt.Scale(scheme="tableau10")),
                    column=alt.Column("type:N", title=None,
                                      header=alt.Header(labelFontSize=10))
                )
                .properties(width=120, height=140)
                .configure_axis(labelFontSize=9, titleFontSize=10)
                .configure_legend(labelFontSize=9, titleFontSize=10)
            )
            st.altair_chart(ch, use_container_width=False)
        else:
            st.info("No sales in the current month yet.")

    with summary_col:
        st.metric("Total Agency Premium (Month)", f"${total_agency:,.0f}")
        if not current_month_sales.empty:
            for _, row in merged.iterrows():
                st.write(f"**{row['agent_name'] or 'Unknown'}**: ${row['sum']:,.0f}  \n↳ Forecast (weekday pace): ${row['forecast']:,.0f}")

    with st.form("sales_form"):
        col1, col2 = st.columns(2)
        with col1:
            customer_name = st.text_input("Customer Name", value="", placeholder="Enter customer name")
            policy_number = st.text_input("Policy Number", value="", placeholder="Enter policy number")
            premium = st.number_input("Premium ($)", min_value=0.0, step=10.0)
            carrier = st.text_input("Carrier", value="", placeholder="Enter carrier")
        with col2:
            effective_date = st.date_input("Effective Date", value=datetime.today())
            date_of_sale = st.date_input("Date of Sale", value=datetime.today())
            agent_name = st.text_input("Agent Name", value="", placeholder="Enter agent name")
            notes = st.text_area("Notes", value="", placeholder="Optional notes")
        submitted = st.form_submit_button("Submit Sale")
        if submitted:
            try:
                supabase.table("daily_sales").insert({
                    "customer_name": customer_name.strip(),
                    "policy_number": policy_number.strip(),
                    "premium": str(premium),
                    "effective_date": str(effective_date),
                    "date_of_sale": str(date_of_sale),
                    "agent_name": agent_name.strip(),
                    "carrier": carrier.strip(),
                    "notes": notes.strip(),
                    "status": "Active"
                }).execute()
                st.success("Sale submitted successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"Failed to save entry: {e}")

    if not current_month_sales.empty:
        current_month_sales["date_of_sale"] = current_month_sales["date_of_sale"].dt.strftime("%m/%d/%Y")
        current_month_sales["effective_date"] = current_month_sales["effective_date"].dt.strftime("%m/%d/%Y")
        edited_month = st.data_editor(current_month_sales, num_rows="dynamic", use_container_width=True, key="edit_current")
        if st.button("Save Current Month Changes"):
            for idx in edited_month.index:
                row = edited_month.loc[idx]
                supabase.table("daily_sales").update({
                    "customer_name": row["customer_name"],
                    "policy_number": row["policy_number"],
                    "premium": row["premium"],
                    "effective_date": str(pd.to_datetime(row["effective_date"]).date()),
                    "date_of_sale": str(pd.to_datetime(row["date_of_sale"]).date()),
                    "agent_name": row["agent_name"],
                    "carrier": row["carrier"],
                    "notes": row["notes"],
                    "status": row["status"]
                }).eq("id", row["id"]).execute()
            st.success("Changes saved successfully.")
            st.rerun()

elif menu == "2. Search and Edit Sales with Date Range":
    st.subheader("Search and Edit All Sales")

    if "search_term" not in st.session_state:
        st.session_state["search_term"] = ""

    wrap = st.container()
    in_col, x_col = wrap.columns([12, 1])

    with x_col:
        st.write("")
        st.write("")
        if st.button("✕", key="clear_search_btn", help="Clear search"):
            st.session_state["search_term"] = ""
            st.rerun()

    with in_col:
        st.text_input(
            "Search by Customer Name or Policy Number",
            key="search_term",
            placeholder="Type to search"
        )

    term = st.session_state["search_term"].lower().strip()

    search_results = (
        df_sales[df_sales.apply(
            lambda r: term in str(r.get("customer_name", "")).lower() or
                      term in str(r.get("policy_number", "")).lower(), axis=1)]
        if term else df_sales.copy()
    )

    if not search_results.empty:
        search_results["date_of_sale"] = pd.to_datetime(search_results["date_of_sale"], errors="coerce").dt.strftime("%m/%d/%Y")
        search_results["effective_date"] = pd.to_datetime(search_results["effective_date"], errors="coerce").dt.strftime("%m/%d/%Y")

    edited_search = st.data_editor(search_results, num_rows="dynamic", use_container_width=True, key="edit_all")
    if st.button("Save All Sales Changes"):
        for idx in edited_search.index:
            row = edited_search.loc[idx]
            supabase.table("daily_sales").update({
                "customer_name": row["customer_name"],
                "policy_number": row["policy_number"],
                "premium": row["premium"],
                "effective_date": str(pd.to_datetime(row["effective_date"]).date()),
                "date_of_sale": str(pd.to_datetime(row["date_of_sale"]).date()),
                "agent_name": row["agent_name"],
                "carrier": row["carrier"],
                "notes": row["notes"],
                "status": row["status"]
            }).eq("id", row["id"]).execute()
        st.success("Changes saved successfully.")
        st.rerun()

    st.subheader("Sales by Date Range")
    date_range = st.date_input("Select Date Range", [])
    if len(date_range) == 2:
        start_date, end_date = date_range
        range_df = df_sales[
            (df_sales["date_of_sale"] >= pd.to_datetime(start_date)) &
            (df_sales["date_of_sale"] <= pd.to_datetime(end_date))
        ].copy()
        range_df["date_of_sale"] = range_df["date_of_sale"].dt.strftime("%m/%d/%Y")
        range_df["effective_date"] = range_df["effective_date"].dt.strftime("%m/%d/%Y")
        st.dataframe(range_df, use_container_width=True)
        buffer = BytesIO()
        range_df.to_excel(buffer, index=False)
        st.download_button("Download Report", buffer.getvalue(), file_name="sales_report.xlsx")

elif menu == "3. Upload Commission, Detect Duplicates, Upload History":
    st.subheader("Upload Commission Statement")
    commission_file = st.file_uploader("Upload Commission File (Policy Numbers Only)", type=["xlsx"])
    if commission_file:
        try:
            commission_df = pd.read_excel(commission_file)
            commission_df["policy_number"] = commission_df["policy_number"].astype(str).str.strip()
            df_sales["policy_number"] = df_sales["policy_number"].astype(str).str.strip()
            df_sales["Matched"] = df_sales["policy_number"].isin(commission_df["policy_number"])
            matched_df = df_sales[df_sales["Matched"] == True]
            unmatched_df = df_sales[df_sales["Matched"] == False]
            st.success(f"Matched: {len(matched_df)}, Unmatched: {len(unmatched_df)}")
            st.download_button("Download Matched", matched_df.to_csv(index=False), "matched_policies.csv")
            st.download_button("Download Unmatched", unmatched_df.to_csv(index=False), "unmatched_policies.csv")
        except Exception as e:
            st.error(f"Error processing file: {e}")

    st.subheader("Duplicate Detection")
    _resp = supabase.table("daily_sales").select("*").execute()
    _df = pd.DataFrame(_resp.data)
    if not _df.empty:
        _df["policy_number"] = _df["policy_number"].astype(str).str.strip()
        duplicate_df = _df[_df.duplicated(subset="policy_number", keep=False)].copy()
    else:
        duplicate_df = pd.DataFrame()

    if not duplicate_df.empty:
        duplicate_df = duplicate_df.sort_values(["policy_number", "id"]).reset_index(drop=True)
        duplicate_df["Select"] = False
        original_dupes = duplicate_df.copy()

        edited_dupes = st.data_editor(
            duplicate_df,
            num_rows="dynamic",
            use_container_width=True,
            key="duplicates_editor"
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("Save Duplicate Edits"):
                exclude_cols = {"Select"}
                changed_ids = []
                orig_indexed = original_dupes.set_index("id")
                edit_indexed = edited_dupes.set_index("id")
                common_cols = [c for c in edit_indexed.columns if c not in exclude_cols]

                for _id in edit_indexed.index:
                    if _id in orig_indexed.index:
                        if not edit_indexed.loc[_id, common_cols].equals(orig_indexed.loc[_id, common_cols]):
                            changed_ids.append(_id)

                for _id in changed_ids:
                    r = edit_indexed.loc[_id]
                    try:
                        supabase.table("daily_sales").update({
                            "customer_name": r.get("customer_name", ""),
                            "policy_number": str(r.get("policy_number", "")).strip(),
                            "premium": r.get("premium"),
                            "effective_date": str(pd.to_datetime(r.get("effective_date", date.today())).date())
                                if pd.notna(r.get("effective_date")) else None,
                            "date_of_sale": str(pd.to_datetime(r.get("date_of_sale", date.today())).date())
                                if pd.notna(r.get("date_of_sale")) else None,
                            "agent_name": r.get("agent_name", ""),
                            "carrier": r.get("carrier", ""),
                            "notes": r.get("notes", ""),
                            "status": r.get("status", "Active")
                        }).eq("id", int(_id)).execute()
                    except Exception as e:
                        st.error(f"Failed to update id={_id}: {e}")
                if changed_ids:
                    st.success(f"Saved {len(changed_ids)} edit(s). Refreshing…")
                    st.rerun()
                else:
                    st.info("No changes detected.")

        with c2:
            if st.button("Delete Selected Duplicates"):
                selected_ids = edited_dupes[edited_dupes["Select"] == True]["id"].tolist()
                for _id in selected_ids:
                    try:
                        supabase.table("daily_sales").delete().eq("id", int(_id)).execute()
                    except Exception as e:
                        st.error(f"Failed to delete id={_id}: {e}")
                st.success(f"Deleted {len(selected_ids)} duplicate row(s). Refreshing…")
                st.rerun()
    else:
        st.info("No duplicate policies found.")

    st.subheader("Upload Historical Daily Sales")
    hist_file = st.file_uploader("Upload Historical Sales (.xlsx)", type=["xlsx"])
    if hist_file:
        try:
            hist_df = pd.read_excel(hist_file)
            hist_df["policy_number"] = hist_df["policy_number"].astype(str).str.strip()
            existing_policies = _df["policy_number"].astype(str).str.strip().tolist() if not _df.empty else []
            new_entries = hist_df[~hist_df["policy_number"].isin(existing_policies)].copy()
            for _, row in new_entries.iterrows():
                supabase.table("daily_sales").insert({
                    "customer_name": str(row.get("customer_name", "")).strip(),
                    "policy_number": str(row.get("policy_number", "")).strip(),
                    "premium": str(row.get("premium", "0")),
                    "effective_date": str(pd.to_datetime(row.get("effective_date", date.today())).date()),
                    "date_of_sale": str(pd.to_datetime(row.get("date_of_sale", date.today())).date()),
                    "agent_name": str(row.get("agent_name", "")).strip(),
                    "carrier": str(row.get("carrier", "")).strip(),
                    "notes": str(row.get("notes", "")).strip(),
                    "status": str(row.get("status", "Active")).strip()
                }).execute()
            st.success(f"{len(new_entries)} new historical sales uploaded.")
            backup_path = os.path.join(os.getcwd(), f"backup_historical_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            hist_df.to_excel(backup_path, index=False)
        except Exception as e:
            st.error(f"Upload failed: {e}")

elif menu == "4. Yearly Sales Chart":
    if df_sales.empty:
        st.info("No sales data available.")
    else:
        years = sorted(df_sales["year"].dropna().unique().astype(int))
        selected_year = st.selectbox("Select Year", years, index=len(years) - 1 if years else 0)
        agents = sorted(df_sales["agent_name"].dropna().unique())
        selected_agents = st.multiselect("Select Agents", agents, default=agents)

        year_df = df_sales[(df_sales["year"] == selected_year) & (df_sales["agent_name"].isin(selected_agents))].copy()
        if year_df.empty:
            st.info("No sales for the selected year/agents.")
        else:
            year_df["month_abbr"] = year_df["month"].apply(lambda m: calendar.month_abbr[int(m)] if pd.notnull(m) else "")
            monthly = (
                year_df.groupby(["month_abbr", "agent_name"], as_index=False)["premium"]
                .sum()
            )

            month_order = [calendar.month_abbr[i] for i in range(1, 13)]
            ch_year = (
                alt.Chart(monthly)
                .mark_bar()
                .encode(
                    x=alt.X("month_abbr:N", sort=month_order, title="Month",
                            axis=alt.Axis(labelAngle=0, labelFontSize=9, titleFontSize=10)),
                    xOffset=alt.X("agent_name:N"),
                    y=alt.Y("premium:Q", title="Premium"),
                    color=alt.Color("agent_name:N", title="Agent", scale=alt.Scale(scheme="tableau10")),
                    tooltip=["month_abbr", "agent_name", alt.Tooltip("premium:Q", format=",.0f")]
                )
                .properties(width=28, height=160, title=f"Monthly Sales by Agent — {selected_year}")
                .configure_axis(labelFontSize=9, titleFontSize=10)
                .configure_legend(labelFontSize=9, titleFontSize=10)
            )
            st.altair_chart(ch_year, use_container_width=True)

            st.metric("Total Agency Premium", f"${year_df['premium'].sum():,.0f}")
