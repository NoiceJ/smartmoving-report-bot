import pandas as pd
import numpy as np
from datetime import datetime
from weasyprint import HTML
import os

# Non-target / Disqualified reasons based on your business rules
NON_TARGET_REASONS = [
    "Problematic Client",
    "Can't Service",
    "Spam / Duplicate / Wrong Call",
    "Last Minute Request",
    "Shopping Around"
]

def clean_currency(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    return float(str(val).replace('$', '').replace(',', '').strip() or 0.0)

def process_smartmoving_reports(files_dict, report_type="weekly", period_label="10/26 - 11/01"):
    """
    files_dict expects paths to:
    - 'lead_source': lead-source-and-conversion-analysis.xlsx
    - 'lost_leads': lost-leads-opportunities-details.xlsx
    - 'new_leads': new-leads.xlsx
    - 'sales_person': sales-person-performance.xlsx
    - 'cancellations': cancellation-summary.xlsx
    """
    
    # 1. Load DataFrames (handling SmartMoving header offsets automatically)
    def load_sm_excel(path):
        xl = pd.ExcelFile(path)
        df = xl.parse(xl.sheet_names[0])
        # Find the actual header row if title rows exist at the top
        for i, row in df.iterrows():
            if row.notna().sum() >= 3:
                df.columns = df.iloc[i]
                df = df.iloc[i+1:].reset_index(drop=True)
                break
        df.columns = [str(c).strip() for c in df.columns]
        return df

    df_sources = load_sm_excel(files_dict['lead_source'])
    df_lost = load_sm_excel(files_dict['lost_leads'])
    df_sales = load_sm_excel(files_dict['sales_person'])
    df_cancel = load_sm_excel(files_dict['cancellations'])

    # Filter out "Total" summary rows at the bottom of SmartMoving exports
    df_sources = df_sources[~df_sources.iloc[:, 0].astype(str).str.contains('Total|Grand Total', case=False, na=False)]
    df_sales = df_sales[~df_sales.iloc[:, 0].astype(str).str.contains('Total|Grand Total', case=False, na=False)]
    df_lost = df_lost[~df_lost.iloc[:, 0].astype(str).str.contains('Total|Grand Total', case=False, na=False)]

    # 2. Core Metrics Calculation
    # Adjust column names dynamically to match SmartMoving standard export headers
    leads_col = [c for c in df_sources.columns if 'Lead' in c or 'Count' in c or 'Received' in c][0]
    quotes_col = [c for c in df_sources.columns if 'Quote' in c][0]
    booked_col = [c for c in df_sources.columns if 'Book' in c][0]
    rev_col = [c for c in df_sales.columns if 'Total' in c or 'Revenue' in c or 'Booked' in c][-1]

    leads_received = int(pd.to_numeric(df_sources[leads_col], errors='coerce').sum())
    quotes_sent = int(pd.to_numeric(df_sources[quotes_col], errors='coerce').sum())
    booked_jobs = int(pd.to_numeric(df_sources[booked_col], errors='coerce').sum())
    booked_total_rev = df_sales[rev_col].apply(clean_currency).sum()
    avg_check = (booked_total_rev / booked_jobs) if booked_jobs > 0 else 0.0

    # Lost & Cancellations
    total_lost = len(df_lost)
    cancelled_count = len(df_cancel[~df_cancel.iloc[:, 0].astype(str).str.contains('Total', case=False, na=False)])

    # Lost Reasons Breakdown
    reason_col = [c for c in df_lost.columns if 'Reason' in c][0]
    lost_counts = df_lost[reason_col].value_counts()

    # Disqualified / Non-Target Leads
    disqualified_count = df_lost[df_lost[reason_col].isin(NON_TARGET_REASONS)].shape[0]
    valid_leads = max(leads_received - disqualified_count, 1)

    # Conversion Rates
    overall_conv = (booked_jobs / leads_received * 100) if leads_received > 0 else 0
    qualified_conv = (booked_jobs / valid_leads * 100) if valid_leads > 0 else 0
    valid_pct = (valid_leads / leads_received * 100) if leads_received > 0 else 0
    quote_to_booked = (booked_jobs / quotes_sent * 100) if quotes_sent > 0 else 0

    # 3. Generate Weekly Text Output
    lost_lines = []
    for reason, count in lost_counts.head(6).items():
        pct = (count / total_lost * 100) if total_lost > 0 else 0
        lost_lines.append(f"{reason} - {pct:.2f}% ({count})")

    weekly_text = f"""Weekly Performance Report {period_label}
Leads Received: {leads_received}
Lost: {total_lost} ({(total_lost/leads_received*100):.0f}%)
Cancelled: {cancelled_count} ({(cancelled_count/leads_received*100):.0f}%)
Disqualified leads / Bad: {disqualified_count} ({(disqualified_count/leads_received*100):.2f}%)
Quotes Sent: {quotes_sent} ({(quotes_sent/leads_received*100):.0f}%)
Booked: {booked_jobs} ({overall_conv:.0f}%)
Conversion: {overall_conv:.0f}%
Qualified Conversion {valid_leads} out of {leads_received} leads were valid ({valid_pct:.1f}%)
Quote -> Booked Conversion: {quote_to_booked:.0f}%
Booked Total: ${booked_total_rev:,.2f}
Main Lost Reasons:
""" + "\n".join(lost_lines)

    # 4. Generate Monthly PDF Output
    source_name_col = df_sources.columns[0]
    sources_html = "".join([
        f"<li><strong>{row[source_name_col]}:</strong> {int(row[leads_col])}</li>"
        for _, row in df_sources.iterrows() if pd.notna(row[source_name_col])
    ])

    all_lost_html = "".join([
        f"<li><strong>{reason}:</strong> {(count/total_lost*100):.2f}% — {count} leads</li>"
        for reason, count in lost_counts.items()
    ])

    agent_name_col = df_sales.columns[0]
    agents_html = "".join([
        f"<tr><td>{row[agent_name_col]}</td><td>{row[leads_col]}</td><td>{row[quotes_col]}</td><td>{row[booked_col]}</td><td>${clean_currency(row[rev_col]):,.2f}</td></tr>"
        for _, row in df_sales.iterrows() if pd.notna(row[agent_name_col])
    ])

    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: #111; padding: 40px; }}
            .cover {{ background: #000; color: #fff; padding: 60px 40px; border-radius: 8px; margin-bottom: 30px; }}
            .cover h1 {{ font-size: 36px; margin: 0; }}
            .cover p {{ color: #ccc; font-size: 18px; }}
            .grid {{ display: flex; justify-content: space-between; margin-bottom: 30px; }}
            .card {{ background: #f4f4f4; padding: 20px; border-radius: 6px; width: 22%; text-align: center; }}
            .card h2 {{ margin: 0; font-size: 24px; color: #000; }}
            .card span {{ font-size: 12px; color: #555; }}
            h3 {{ border-bottom: 2px solid #000; padding-bottom: 6px; margin-top: 30px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            th, td {{ border: 1px solid #ddd; padding: 10px; text-align: left; }}
            th {{ background: #000; color: #fff; }}
            ul {{ columns: 2; }}
        </style>
    </head>
    <body>
        <div class="cover">
            <h1>We Sell — Monthly Sales Performance Report</h1>
            <p>{period_label}</p>
        </div>

        <h3>01. Financial Overview</h3>
        <div class="grid">
            <div class="card"><h2>${booked_total_rev:,.2f}</h2><span>Booked Total</span></div>
            <div class="card"><h2>${avg_check:,.2f}</h2><span>Average Check</span></div>
            <div class="card"><h2>{booked_jobs}</h2><span>Booked Jobs</span></div>
            <div class="card"><h2>{leads_received}</h2><span>Total Leads</span></div>
        </div>

        <h3>02. Conversion Metrics</h3>
        <p><strong>Overall Conversion Rate:</strong> {overall_conv:.2f}%</p>
        <p><strong>Qualified Leads Conversion Rate:</strong> {qualified_conv:.2f}% ({valid_leads} target leads)</p>
        <p><strong>Quote-to-Booked Conversion Rate:</strong> {quote_to_booked:.2f}%</p>

        <h3>03. Lead Sources</h3>
        <ul>{sources_html}</ul>

        <h3>04. Lost Leads Analysis (Total: {total_lost} leads)</h3>
        <ul>{all_lost_html}</ul>

        <h3>05. Agent Statistics</h3>
        <table>
            <tr><th>Sales Rep</th><th>Leads</th><th>Quotes</th><th>Booked</th><th>Booked Total</th></tr>
            {agents_html}
        </table>
    </body>
    </html>
    """

    pdf_path = f"/tmp/Monthly_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    HTML(string=html_content).write_pdf(pdf_path)

    return {
        "weekly_text": weekly_text,
        "pdf_path": pdf_path
    }
