"""
Automated Certificate Generation and Email Sending System
DV Analytics

Run with:  streamlit run app.py
"""

import os
import io
import zipfile

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from certificate_generator import (
    load_template_as_image,
    generate_certificate,
    render_certificate_image,
    suggest_text_style,
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
# Page config & modern theme
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="DV Analytics — Certificate & Email Suite",
    page_icon="🎓",
    layout="wide",
)

NAVY = "#0B1B4D"
NAVY_DEEP = "#060F30"
RED = "#EF233C"
ACCENT = "#5B6CF7"
BG = "#F3F4FA"
CARD = "#FFFFFF"
BORDER = "#E6E8F5"
MUTED = "#6B7188"

st.markdown(
    f"""
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@500;600;700;800&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
        .stApp {{ background: {BG}; }}
        #MainMenu, footer {{ visibility: hidden; }}

        .dv-hero {{
            background: radial-gradient(120% 160% at 0% 0%, {ACCENT}33 0%, transparent 45%),
                        linear-gradient(120deg, {NAVY} 0%, {NAVY_DEEP} 100%);
            padding: 30px 34px; border-radius: 20px; margin-bottom: 26px;
            display: flex; align-items: center; gap: 18px;
            box-shadow: 0 12px 30px -12px rgba(11,27,77,0.55);
        }}
        .dv-hero .mark {{
            width: 52px; height: 52px; border-radius: 14px;
            background: linear-gradient(135deg, {RED}, #ff6b6b);
            display: flex; align-items: center; justify-content: center;
            color: white; font-weight: 800; font-size: 19px;
            font-family: 'Poppins', sans-serif;
            box-shadow: 0 6px 16px -4px {RED}aa;
        }}
        .dv-hero h1 {{
            color: white; font-size: 23px; margin: 0; font-family: 'Poppins', sans-serif; font-weight: 700;
        }}
        .dv-hero p {{ color: #B7BEEF; font-size: 13px; margin: 3px 0 0; }}

        .dv-card {{
            background: {CARD}; border: 1px solid {BORDER}; border-radius: 16px;
            padding: 22px 24px; margin-bottom: 18px;
            box-shadow: 0 2px 10px -6px rgba(11,27,77,0.08);
        }}
        .dv-section-title {{
            font-family: 'Poppins', sans-serif; font-weight: 700; font-size: 17px;
            color: {NAVY}; margin-bottom: 2px;
        }}
        .dv-section-sub {{ color: {MUTED}; font-size: 13px; margin-bottom: 14px; }}

        div[data-testid="stTabs"] button[data-baseweb="tab"] {{
            font-family: 'Poppins', sans-serif; font-weight: 600; font-size: 14px;
            color: {MUTED}; padding: 10px 6px;
        }}
        div[data-testid="stTabs"] button[aria-selected="true"] {{ color: {NAVY}; }}
        div[data-testid="stTabs"] div[data-baseweb="tab-highlight"] {{
            background: linear-gradient(90deg, {ACCENT}, {RED}); height: 3px; border-radius: 3px;
        }}

        div[data-testid="stMetric"] {{
            background: {CARD}; border: 1px solid {BORDER}; border-radius: 14px;
            padding: 14px 18px; box-shadow: 0 2px 8px -6px rgba(11,27,77,0.08);
        }}
        div[data-testid="stMetricLabel"] {{ color: {MUTED}; font-weight: 500; }}
        div[data-testid="stMetricValue"] {{ color: {NAVY}; font-family: 'Poppins', sans-serif; }}

        .stButton>button {{
            background: linear-gradient(135deg, {NAVY} 0%, {ACCENT} 130%);
            color: white; border-radius: 10px; font-weight: 600; border: none;
            padding: 0.55em 1.3em; font-family: 'Poppins', sans-serif; font-size: 14px;
            box-shadow: 0 6px 16px -6px {NAVY}88; transition: all 0.15s ease;
        }}
        .stButton>button:hover {{
            transform: translateY(-1px); box-shadow: 0 10px 22px -8px {NAVY}aa; color: white;
        }}
        .stButton>button:disabled {{
            background: #D8DAE8; color: #9296AC; box-shadow: none; transform: none;
        }}
        .stDownloadButton>button {{
            background: white; color: {NAVY}; border: 1.5px solid {NAVY}33; border-radius: 10px;
            font-weight: 600; font-family: 'Poppins', sans-serif; font-size: 13.5px;
        }}
        .stDownloadButton>button:hover {{ border-color: {ACCENT}; color: {ACCENT}; }}

        .dv-badge {{
            display: inline-block; padding: 3px 12px; border-radius: 999px;
            font-size: 12px; font-weight: 600; font-family: 'Poppins', sans-serif;
        }}
        .dv-badge-ok {{ background: #E4F7EC; color: #128A44; }}
        .dv-badge-warn {{ background: #FDECEC; color: {RED}; }}

        div[data-testid="stProgress"] > div > div {{
            background: linear-gradient(90deg, {ACCENT}, {RED});
        }}
    </style>

    <div class="dv-hero">
        <div class="mark">DV</div>
        <div>
            <h1>Certificate &amp; Email Suite</h1>
            <p>Automated certificate generation and personalized email delivery</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


def card_start(title: str, subtitle: str = ""):
    sub_html = f'<div class="dv-section-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="dv-card"><div class="dv-section-title">{title}</div>{sub_html}',
        unsafe_allow_html=True,
    )


def card_end():
    st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
defaults = {
    "records": [],
    "column_map": {},
    "template_path": None,
    "template_fingerprint": None,
    "cert_paths": {},
    "results": [],
    "font_size": 60,
    "y_pos_pct": 50,
    "text_color_hex": "#0B1B4D",
    "confirmed_params": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------------------------------------------------------------------------
# Sidebar — SMTP status
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("#### ✉️ SMTP status")
    cfg = SMTPConfig()
    if cfg.is_configured():
        st.markdown('<span class="dv-badge dv-badge-ok">● Connected</span>', unsafe_allow_html=True)
        st.caption(f"{cfg.username}\nvia {cfg.host}:{cfg.port}")
    else:
        st.markdown('<span class="dv-badge dv-badge-warn">● Not configured</span>', unsafe_allow_html=True)
        st.caption(
            "Set **SMTP_EMAIL** and **SMTP_PASSWORD** as environment variables / "
            "Streamlit secrets. Gmail requires an **App Password**, not your normal login."
        )
    st.divider()
    st.caption("DV Analytics · Certificate & Email Suite")

tab1, tab2, tab3, tab4 = st.tabs(
    ["①  Upload", "②  Generate", "③  Send", "④  Dashboard"]
)

# ---------------------------------------------------------------------------
# TAB 1 — Upload
# ---------------------------------------------------------------------------
with tab1:
    card_start("Participant list", "Required fields: Name, Mobile Number, Email ID — headers are matched flexibly.")
    excel_file = st.file_uploader("Excel file (.xlsx)", type=["xlsx"], label_visibility="collapsed")

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
                mapping_str = " · ".join(f"{k} → `{v}`" for k, v in column_map.items())
                st.markdown(f'<span class="dv-badge dv-badge-ok">✓ {len(records)} records validated</span>', unsafe_allow_html=True)
                st.caption(mapping_str)
                st.dataframe(pd.DataFrame(records), use_container_width=True, height=220)
        except Exception as e:
            st.error(f"Could not read Excel file: {e}")
    card_end()

    card_start("Certificate template", "PNG, JPG, or PDF — the name is drawn centered on top of this image.")
    template_file = st.file_uploader("Certificate template", type=["png", "jpg", "jpeg", "pdf"], label_visibility="collapsed")

    if template_file:
        template_path = os.path.join(OUTPUT_DIR, f"template{os.path.splitext(template_file.name)[1]}")
        with open(template_path, "wb") as f:
            f.write(template_file.getbuffer())
        st.session_state.template_path = template_path

        fingerprint = (template_file.name, template_file.size)
        is_new_template = fingerprint != st.session_state.template_fingerprint

        template_img = load_template_as_image(template_path)

        if is_new_template:
            st.session_state.template_fingerprint = fingerprint
            suggested_size, suggested_color = suggest_text_style(
                template_img, st.session_state.y_pos_pct / 100.0
            )
            st.session_state.font_size = suggested_size
            st.session_state.text_color_hex = "#%02x%02x%02x" % suggested_color
            st.session_state.confirmed_params = None
            st.markdown('<span class="dv-badge dv-badge-ok">✓ Template uploaded — style auto-tuned</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="dv-badge dv-badge-ok">✓ Template uploaded</span>', unsafe_allow_html=True)

        with st.expander("🎨 Customize name placement, size & color", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                font_size = st.slider("Font size (px)", 20, 300, key="font_size")
                y_pos = st.slider("Vertical position (% down)", 0, 100, key="y_pos_pct") / 100.0
            with c2:
                color_hex = st.color_picker("Text color", key="text_color_hex")
                if st.button("↺ Auto-fit to this template"):
                    s_size, s_color = suggest_text_style(template_img, st.session_state.y_pos_pct / 100.0)
                    st.session_state.font_size = s_size
                    st.session_state.text_color_hex = "#%02x%02x%02x" % s_color
                    st.session_state.confirmed_params = None
                    st.rerun()

        font_size = st.session_state.font_size
        y_pos = st.session_state.y_pos_pct / 100.0
        color_hex = st.session_state.text_color_hex
        text_color = tuple(int(color_hex.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))

        if st.session_state.records:
            sample_name = st.session_state.records[0]["Name"]
            preview_img = render_certificate_image(
                template_img, sample_name, None, font_size, text_color, y_pos
            )
            st.image(preview_img, caption=f"Live preview — {sample_name}", use_container_width=True)

            current_params = (fingerprint, font_size, round(y_pos, 3), color_hex)
            confirmed = st.checkbox(
                "✅ I can clearly see the name above on the certificate",
                value=(st.session_state.confirmed_params == current_params),
            )
            if confirmed:
                st.session_state.confirmed_params = current_params
            elif st.session_state.confirmed_params == current_params:
                st.session_state.confirmed_params = None
        else:
            st.info("Upload the participant list above to preview a sample name on this template.")
    card_end()

# ---------------------------------------------------------------------------
# TAB 2 — Generate Certificates
# ---------------------------------------------------------------------------
with tab2:
    card_start("Generate certificates")

    records = st.session_state.records
    template_path = st.session_state.template_path

    c1, c2, c3 = st.columns(3)
    c1.metric("Total records", len(records))
    c2.metric("Template ready", "Yes" if template_path else "No")
    c3.metric("Certificates generated", len(st.session_state.cert_paths))

    fingerprint = st.session_state.template_fingerprint
    font_size = st.session_state.font_size
    y_pos = st.session_state.y_pos_pct / 100.0
    color_hex = st.session_state.text_color_hex
    text_color = tuple(int(color_hex.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    current_params = (fingerprint, font_size, round(y_pos, 3), color_hex)
    preview_confirmed = st.session_state.confirmed_params == current_params and fingerprint is not None

    disabled = not (records and template_path)
    if disabled:
        st.info("Upload both the participant list and the certificate template in Step ① first.")
    elif not preview_confirmed:
        st.warning("Go back to Step ① and confirm the name preview looks correct before generating in bulk.")

    if st.button("🎓  Generate Certificates", disabled=disabled or not preview_confirmed, key="btn_generate"):
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
        progress.empty()
        ok_count = sum(1 for v in cert_paths.values() if v)
        st.success(f"Done — {ok_count} of {len(records)} certificates generated.")

    if st.session_state.cert_paths:
        st.markdown("**Review**")
        review_df = pd.DataFrame(
            [{"Name": n, "Certificate Generated": "Yes" if p else "No"}
             for n, p in st.session_state.cert_paths.items()]
        )
        st.dataframe(review_df, use_container_width=True, height=220)

        sample_paths = [p for p in st.session_state.cert_paths.values() if p]
        if sample_paths:
            with st.expander("🔍 Preview a generated certificate"):
                st.image(load_template_as_image(sample_paths[0]), use_container_width=True)
    card_end()

# ---------------------------------------------------------------------------
# TAB 3 — Send Certificates
# ---------------------------------------------------------------------------
with tab3:
    card_start("Compose email")

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
        height=200,
    )
    card_end()

    card_start("Send")
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
        st.info("Generate certificates in Step ② before sending.")
    elif not SMTPConfig().is_configured():
        st.warning("SMTP is not configured — see the sidebar.")

    if st.button("✉️  Send Certificates", disabled=not can_send, key="btn_send"):
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
        progress.empty()

        report_df = build_report_dataframe(results)
        report_df.to_excel(REPORT_PATH, index=False)
        errors_df = report_df[report_df["Error Message"] != ""]
        errors_df.to_excel(ERROR_REPORT_PATH, index=False)

        n_sent = sum(1 for r in results if r["Email Sent"] == "Yes")
        st.success(f"Done — {n_sent} of {total} emails sent successfully.")
        st.rerun()
    card_end()

# ---------------------------------------------------------------------------
# TAB 4 — Report & Dashboard
# ---------------------------------------------------------------------------
with tab4:
    card_start("Dashboard")

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
    card_end()

    card_start("Downloads")
    dl1, dl2, dl3, dl4 = st.columns(4)

    with dl1:
        if cert_paths and any(cert_paths.values()):
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, path in cert_paths.items():
                    if path and os.path.exists(path):
                        zf.write(path, arcname=os.path.basename(path))
            st.download_button("⬇️ Certificates (.zip)", buf.getvalue(), file_name="Certificates.zip", mime="application/zip")
        else:
            st.button("⬇️ Certificates (.zip)", disabled=True)

    with dl2:
        if os.path.exists(REPORT_PATH):
            with open(REPORT_PATH, "rb") as f:
                st.download_button("⬇️ Email Report", f.read(), file_name="Email_Sending_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.button("⬇️ Email Report", disabled=True)

    with dl3:
        if os.path.exists(ERROR_REPORT_PATH):
            with open(ERROR_REPORT_PATH, "rb") as f:
                st.download_button("⬇️ Error Report", f.read(), file_name="Error_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.button("⬇️ Error Report", disabled=True)

    with dl4:
        if os.path.exists(LOG_PATH):
            with open(LOG_PATH, "rb") as f:
                st.download_button("⬇️ Log File", f.read(), file_name="email_log.txt")
        else:
            st.button("⬇️ Log File", disabled=True)
    card_end()

    if results:
        card_start("Full report")
        st.dataframe(build_report_dataframe(results), use_container_width=True, height=320)
        card_end()
