# Automated Certificate Generation and Email Sending System

A Streamlit app that turns an Excel participant list + a certificate template
into personalized PDF certificates, then emails each one out automatically —
with a live progress bar, dashboard, and downloadable reports.

## Features

- Upload an `.xlsx` file with **Name**, **Mobile Number**, **Email ID** columns (flexible column matching, order-independent)
- Upload a PNG/JPG/PDF certificate template; name is drawn centered on top (adjustable font size, position, color)
- One-click **Generate Certificates** → saved as `Name_Certificate.pdf` in `output/certificates/`
- One-click **Send Certificates** → personalized email per participant via SMTP, certificate attached, live progress bar + running Total/Sent/Failed/Pending counts
- Email address validation before sending
- Auto-generated `Email_Sending_Report.xlsx` and `Error_Report.xlsx` after each send
- `email_log.txt` — timestamped log of every send attempt
- Dashboard: total participants, certificates generated, emails sent, failed, success rate
- Download buttons for certificates (.zip), report, error report, and log file
- Credentials read only from environment variables — nothing hardcoded

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env
# edit .env with your real SMTP credentials
```

**Gmail users:** enable 2-Step Verification, then create an
[App Password](https://myaccount.google.com/apppasswords) — use that as
`SMTP_PASSWORD`, not your normal Gmail password.

**Outlook/Office365 users:** set `SMTP_HOST=smtp.office365.com` in `.env`.

## Run

```bash
streamlit run app.py
```

Then open the local URL Streamlit prints (usually `http://localhost:8501`).

## Usage flow

1. **① Upload** — upload your Excel list and certificate template; preview the sample certificate.
2. **② Generate Certificates** — click the button; review the generated list.
3. **③ Send Certificates** — write your subject/body (`{{Name}}` is replaced per participant), click Send, watch the progress bar.
4. **④ Report & Dashboard** — view KPIs and download the certificates zip, the Excel report, the error report, and the log file.

## File structure

```
app.py                    # Streamlit UI and orchestration
certificate_generator.py  # PIL-based certificate rendering + PDF export
email_sender.py           # SMTP wrapper, attachment handling, email validation
utils.py                  # Excel validation, logging, report building
requirements.txt
.env.example
```

## Notes

- Certificates and reports are written to `output/` (created automatically).
- For very large batches (thousands of records), sending happens sequentially
  in one connection to stay within SMTP rate limits — expect it to take a
  while; the progress bar and counters update live throughout.
- SQLite persistence (to resume a partially-completed batch) was left out of
  this version for simplicity, per the spec noting it as optional — the
  Excel report + log file already give you a full audit trail. Ask if you'd
  like resume/retry support added.
