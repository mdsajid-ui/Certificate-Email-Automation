"""
Automated Certificate Generation and Email Sending System
DV Analytics

Run with:  streamlit run app.py
"""

import os
import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from certificate_generator import (
    load_template_as_image,
    generate_certificate,
    render_certificate_image,
)
from email_sender import EmailSender, SMTPConfig, is_valid_email
from utils import (
    validate_excel_columns,
    normalize_records,
    build_report_dataframe,
    log_event,
    now_str,
    LOG_PATH,
)

load_dotenv()

OUTPUT_DIR = "output"
CERT_DIR = os.path.join(OUTPUT_DIR, "certificates")
REPORT_PATH = os.path.join(OUTPUT_DIR, "Email_Sending_Report.xlsx")
ERROR_REPORT_PATH = os.path.join(OUTPUT_DIR, "Error_Report.xlsx")

os.makedirs(CERT_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Page config & branding
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="DV Analytics — Certificate & Email Suite",
    page_icon="🎓",
    layout="wide",
)

NAVY = "#050374"
RED = "#EA1313"

st.markdown(
    f"""
    <style>
        .stApp {{ background-color: #f4f5fa; }}
        .dv-header {{
            background: linear-gradient(120deg, {NAVY} 0%, #03025a 100%);
            padding: 22px 32px; border-radius: 12px; margin-bottom: 24px;
            display: flex; align-items: center; gap: 16px;
        }}
        .dv-header .mark {{
            width: 46px; height: 46px; border-radius: 10px; background: {RED};
            display: flex; align-items: center; justify-content: center;
            color: white; font-weight: 700; font-size: 18px;
        }}
        .dv-header h1 {{ color: white; font-size: 20px; margin: 0; }}
        .dv-header p {{ color: #c7c8ee; font-size: 12.5px; margin: 2px 0 0; }}
        div[data-testid="stMetric"] {{
            background: #fbfbfe; border: 1px solid #e3e4ee; border-radius: 10px;
            padding: 12px 16px;
        }}
        .stButton>button {{
            background-color: {NAVY}; color: white; border-radius: 8px; font-weight: 600;
            border: none;
        }}
        .stButton>button:hover {{ background-color: #03025a; color: white; }}
    </style>
    <div class="dv-header">
        <div class="mark">DV</div>
        <div>
            <h1>Certificate &amp; Email Suite</h1>
            <p>Automated certificate generation and email delivery</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
defaults = {
    "records": [],          # list of {Name, Mobile Number, Email ID}
    "column_map": {},
    "template_path": None,
    "cert_paths": {},       # Name -> path
    "results": [],          # final report rows
    "generated": False,
    "sent": False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------------------------------------------------------------------------
# Sidebar — SMTP status & certificate styling
# ---------------------------------------------------------------------------
with st.sidebar:
    st.subheader("SMTP configuration")
    cfg = SMTPConfig()
    if cfg.is_configured():
        st.success(f"Configured: {cfg.username} via {cfg.host}:{cfg.port}")
    else:
        st.error("Not configured")
        st.caption(
            "Set **SMTP_EMAIL** and **SMTP_PASSWORD** (and optionally SMTP_HOST / "
            "SMTP_PORT) as environment variables before sending. See `.env.example`. "
            "For Gmail, use an **App Password**, not your normal login password."
        )

    st.divider()
    st.subheader("Certificate text style")
    font_size = st.slider("Font size", 20, 150, 60)
    y_pos = st.slider("Vertical position (% down the template)", 0, 100, 50) / 100.0
    color_hex = st.color_picker("Text color", "#050374")
    text_color = tuple(int(color_hex.lstrip("#")[i:i+2], 16) for i in (0, 2, 4))

tab1, tab2, tab3, tab4 = st.tabs(
    ["① Upload", "② Generate Certificates", "③ Send Certificates", "④ Report & Dashboard"]
)

# ---------------------------------------------------------------------------
# TAB 1 — Upload
# ---------------------------------------------------------------------------
with tab1:
    st.markdown("### Upload participant list")
    st.caption("Required columns: **Name**, **Mobile Number**, **Email ID** (column order doesn't matter).")

    excel_file = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])

    if excel_file:
        try:
            df = pd.read_excel(excel_file)
            is_valid, column_map, missing = validate_excel_columns(df)
            if not is_valid:
                st.error(f"Missing required column(s): {', '.join(missing)}")
            else:
                records = normalize_records(df, column_map)
                st.session_state.records = records
                st.session_state.column_map = column_map
                st.success(f"✅ {len(records)} records found and validated.")
                st.dataframe(pd.DataFrame(records), use_container_width=True, height=280)
        except Exception as e:
            st.error(f"Could not read Excel file: {e}")

    st.markdown("### Upload certificate template")
    st.caption("PNG, JPG, or PDF. The participant's name is drawn centered on top of this image.")
    template_file = st.file_uploader("Certificate template", type=["png", "jpg", "jpeg", "pdf"])

    if template_file:
        template_path = os.path.join(OUTPUT_DIR, f"template{os.path.splitext(template_file.name)[1]}")
        with open(template_path, "wb") as f:
            f.write(template_file.getbuffer())
        st.session_state.template_path = template_path
        st.success(f"Template uploaded: {template_file.name}")

        if st.session_state.records:
            try:
                sample_name = st.session_state.records[0]["Name"]
                template_img = load_template_as_image(template_path)
                preview_img = render_certificate_image(
                    template_img, sample_name, None, font_size, text_color, y_pos
                )
                st.image(preview_img, caption=f"Preview — {sample_name}", use_container_width=True)
            except Exception as e:
                st.warning(f"Preview unavailable: {e}")

# ---------------------------------------------------------------------------
# TAB 2 — Generate Certificates
# ---------------------------------------------------------------------------
with tab2:
    st.markdown("### Generate certificates")

    records = st.session_state.records
    template_path = st.session_state.template_path

    c1, c2, c3 = st.columns(3)
    c1.metric("Total records", len(records))
    c2.metric("Template ready", "Yes" if template_path else "No")
    c3.metric("Certificates generated", len(st.session_state.cert_paths))

    disabled = not (records and template_path)
    if disabled:
        st.info("Upload both the participant list and the certificate template in Step 1 first.")

    if st.button("🎓 Generate Certificates", disabled=disabled, key="btn_generate"):
        template_img = load_template_as_image(template_path)
        progress = st.progress(0, text="Starting...")
        used_names = {}
        cert_paths = {}
        for i, rec in enumerate(records):
            name = rec["Name"]
            try:
                path = generate_certificate(
                    template_img, name, CERT_DIR,
                    font_path=None, font_size=font_size,
                    text_color=text_color, y_position_pct=y_pos,
                    used_names=used_names,
                )
                cert_paths[name] = path
            except Exception as e:
                cert_paths[name] = None
                log_event(rec.get("Email ID", ""), "CERT_GENERATION_FAILED", str(e))
            progress.progress((i + 1) / len(records), text=f"Generating {i+1} of {len(records)} — {name}")
        st.session_state.cert_paths = cert_paths
        st.session_state.generated = True
        progress.empty()
        ok_count = sum(1 for v in cert_paths.values() if v)
        st.success(f"Done — {ok_count} of {len(records)} certificates generated in `{CERT_DIR}/`.")

    if st.session_state.cert_paths:
        st.markdown("#### Review")
        review_df = pd.DataFrame(
            [{"Name": n, "Certificate Generated": "Yes" if p else "No"}
             for n, p in st.session_state.cert_paths.items()]
        )
        st.dataframe(review_df, use_container_width=True, height=260)

# ---------------------------------------------------------------------------
# TAB 3 — Send Certificates
# ---------------------------------------------------------------------------
with tab3:
    st.markdown("### Compose email")

    subject = st.text_input("Subject", "Congratulations! Your Certificate is Ready")
    body_template = st.text_area(
        "Body (use {{Name}} to insert the participant's name)",
        value=(
            "Dear {{Name}},\n\n"
            "Thank you for participating in our program.\n"
            "Please find your certificate attached to this email.\n\n"
            "We appreciate your participation and wish you all the best for your "
            "future endeavors.\n\n"
            "Regards,\nDV Analytics Team"
        ),
        height=220,
    )

    records = st.session_state.records
    cert_paths = st.session_state.cert_paths

    total = len(records)
    sent_count = sum(1 for r in st.session_state.results if r.get("Email Sent") == "Yes")
    failed_count = sum(1 for r in st.session_state.results if r.get("Email Sent") == "No")
    pending_count = total - sent_count - failed_count

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Records", total)
    m2.metric("Sent Successfully", sent_count)
    m3.metric("Failed", failed_count)
    m4.metric("Pending", max(pending_count, 0))

    can_send = bool(records) and bool(cert_paths) and SMTPConfig().is_configured()
    if not records or not cert_paths:
        st.info("Generate certificates in Step 2 before sending.")
    elif not SMTPConfig().is_configured():
        st.warning("SMTP is not configured — set SMTP_EMAIL / SMTP_PASSWORD env vars (see sidebar).")

    if st.button("✉️ Send Certificates", disabled=not can_send, key="btn_send"):
        progress = st.progress(0, text="Starting...")
        results = []
        sender = EmailSender()
        try:
            sender.connect()
        except Exception as e:
            st.error(f"Could not connect to SMTP server: {e}")
            sender = None

        for i, rec in enumerate(records):
            name = rec["Name"]
            mobile = rec.get("Mobile Number", "")
            email = rec.get("Email ID", "")
            cert_path = cert_paths.get(name)
            row = {
                "Name": name,
                "Mobile Number": mobile,
                "Email ID": email,
                "Certificate Generated": "Yes" if cert_path else "No",
                "Email Sent": "No",
                "Sent Date & Time": "",
                "Error Message": "",
            }

            progress.progress((i + 1) / total, text=f"Sending {i+1} of {total} — {name}")

            if not cert_path:
                row["Error Message"] = "Certificate not generated"
                log_event(email, "SKIPPED", "Certificate not generated")
                results.append(row)
                continue
            if not is_valid_email(email):
                row["Error Message"] = "Invalid email address"
                log_event(email, "FAILED", "Invalid email address")
                results.append(row)
                continue
            if sender is None:
                row["Error Message"] = "SMTP connection unavailable"
                results.append(row)
                continue

            personalized_body = body_template.replace("{{Name}}", name)
            try:
                sender.send(email, subject, personalized_body, cert_path)
                row["Email Sent"] = "Yes"
                row["Sent Date & Time"] = now_str()
                log_event(email, "SENT")
            except Exception as e:
                row["Error Message"] = str(e)
                log_event(email, "FAILED", str(e))

            results.append(row)

        if sender is not None:
            sender.close()

        st.session_state.results = results
        st.session_state.sent = True
        progress.empty()

        report_df = build_report_dataframe(results)
        report_df.to_excel(REPORT_PATH, index=False)
        errors_df = report_df[report_df["Error Message"] != ""]
        errors_df.to_excel(ERROR_REPORT_PATH, index=False)

        n_sent = sum(1 for r in results if r["Email Sent"] == "Yes")
        st.success(f"Done — {n_sent} of {total} emails sent successfully.")
        st.rerun()

# ---------------------------------------------------------------------------
# TAB 4 — Report & Dashboard
# ---------------------------------------------------------------------------
with tab4:
    st.markdown("### Dashboard")

    records = st.session_state.records
    cert_paths = st.session_state.cert_paths
    results = st.session_state.results

    total = len(records)
    certs_generated = sum(1 for v in cert_paths.values() if v)
    emails_sent = sum(1 for r in results if r.get("Email Sent") == "Yes")
    emails_failed = sum(1 for r in results if r.get("Email Sent") == "No")
    success_rate = f"{(emails_sent / total * 100):.1f}%" if total else "0.0%"

    d1, d2, d3, d4, d5 = st.columns(5)
    d1.metric("Total Participants", total)
    d2.metric("Certificates Generated", certs_generated)
    d3.metric("Emails Sent", emails_sent)
    d4.metric("Failed Emails", emails_failed)
    d5.metric("Success Rate", success_rate)

    st.markdown("### Download options")
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        if cert_paths and any(cert_paths.values()):
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, path in cert_paths.items():
                    if path and os.path.exists(path):
                        zf.write(path, arcname=os.path.basename(path))
            st.download_button(
                "⬇️ Download Certificates (.zip)", buf.getvalue(),
                file_name="Certificates.zip", mime="application/zip",
            )
        else:
            st.button("⬇️ Download Certificates (.zip)", disabled=True)

    with dl2:
        if os.path.exists(REPORT_PATH):
            with open(REPORT_PATH, "rb") as f:
                st.download_button(
                    "⬇️ Download Email Report", f.read(),
                    file_name="Email_Sending_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.button("⬇️ Download Email Report", disabled=True)

    with dl3:
        if os.path.exists(ERROR_REPORT_PATH):
            with open(ERROR_REPORT_PATH, "rb") as f:
                st.download_button(
                    "⬇️ Download Error Report", f.read(),
                    file_name="Error_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.button("⬇️ Download Error Report", disabled=True)

    if os.path.exists(LOG_PATH):
        with open(LOG_PATH, "rb") as f:
            st.download_button("⬇️ Download Log File (email_log.txt)", f.read(), file_name="email_log.txt")

    if results:
        st.markdown("### Full report")
        st.dataframe(build_report_dataframe(results), use_container_width=True, height=320)
