"""Generate the synthetic demo content: Trust Center documents, a prospect
questionnaire, and a sample NDA.

Everything here is FICTIONAL sample content for a Box demo. It is not any
real company's security policy.

Run:  python make_content.py      -> writes files into ./demo_content/
"""
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUT = Path(__file__).parent / "demo_content"
DISCLAIMER = "SAMPLE: fictional content created for a Box demo. Not a real company policy."

# (file name, title, review-by date, sections). One document is intentionally
# past its review date so the agent's stale-source check has something to flag.
TRUST_DOCS = [
    ("Data Residency and Hosting (SAMPLE).docx", "Data Residency and Hosting", "2027-03-31", [
        ("Hosting regions", "The platform runs on major public cloud providers. Customers choose a hosting region when the contract is signed: US East, EU (Frankfurt), UK (London), Australia (Sydney) or Japan (Tokyo)."),
        ("EU customers", "Customer telemetry for EU-hosted customers is stored and processed in the EU (Frankfurt) region and does not leave the EU. Support staff outside the EU can only access EU data through approved, logged sessions."),
        ("Data retention", "Telemetry retention is configurable per customer from 10 to 365 days. When a contract ends, customer data is deleted within 30 days and deletion is confirmed in writing on request."),
        ("Subprocessors", "A current list of subprocessors is published on the Trust Center. Customers are notified 30 days before a new subprocessor is added."),
    ]),
    ("Encryption and Key Management (SAMPLE).docx", "Encryption and Key Management", "2027-03-31", [
        ("In transit", "All data in transit is encrypted with TLS 1.2 or higher. Older protocols and weak cipher suites are disabled."),
        ("At rest", "All customer data at rest is encrypted with AES-256."),
        ("Key management", "Encryption keys are held in the cloud provider's key management service. Keys are rotated every 12 months and access to keys is restricted to a small, named operations group."),
        ("Customer-managed keys", "Customer-managed encryption keys are not currently supported."),
    ]),
    ("Identity and Access Management (SAMPLE).docx", "Identity and Access Management", "2027-01-31", [
        ("Employee access", "Employees sign in with single sign-on and multi-factor authentication. Access follows least privilege and role-based access control."),
        ("Production access", "Access to production systems is granted just in time, for a maximum of 8 hours, and requires manager approval. All production sessions are logged."),
        ("Access reviews", "Access rights are reviewed every quarter. Leavers lose all access within 24 hours of their last day."),
        ("Customer access", "Customers can enforce SAML single sign-on for their users and can manage users automatically with SCIM."),
    ]),
    ("Incident Response (SAMPLE).docx", "Incident Response", "2027-02-28", [
        ("Detection and triage", "A security operations team monitors the environment 24 hours a day, 7 days a week. Security incidents are triaged within 1 hour of detection."),
        ("Customer notification", "Affected customers are notified without undue delay and within 72 hours of a confirmed personal data breach."),
        ("Testing", "The incident response plan is tested at least once a year through tabletop exercises, and lessons learned are tracked to closure."),
    ]),
    ("Business Continuity and Disaster Recovery (SAMPLE).docx", "Business Continuity and Disaster Recovery", "2026-06-30", [
        ("Objectives", "The recovery time objective (RTO) is 4 hours and the recovery point objective (RPO) is 1 hour for the core platform."),
        ("Backups", "Backups are taken daily, encrypted, and stored in a separate availability zone within the same region."),
        ("Testing", "Disaster recovery is tested once a year."),
    ]),
    ("Vulnerability Management and Penetration Testing (SAMPLE).docx", "Vulnerability Management and Penetration Testing", "2027-04-30", [
        ("Scanning", "Infrastructure and container images are scanned for vulnerabilities continuously. Critical vulnerabilities are remediated within 7 days and high vulnerabilities within 30 days."),
        ("Penetration testing", "An independent third party performs a penetration test at least once a year. An executive summary of the latest test is available to customers under NDA."),
        ("Secure development", "Code changes require peer review, and automated static analysis and dependency scanning run on every build."),
    ]),
    ("AI and Customer Data Usage (SAMPLE).docx", "AI and Customer Data Usage", "2027-05-31", [
        ("Model training", "Customer data is not used to train generative AI models shared across customers."),
        ("AI governance", "AI features are governed under an AI management system aligned with ISO 42001, including risk assessments for each new AI capability."),
        ("Customer control", "Customers can turn generative AI features on or off for their environment."),
    ]),
    ("Compliance and Certifications Overview (SAMPLE).docx", "Compliance and Certifications Overview", "2027-06-30", [
        ("Certifications", "The platform holds SOC 2 Type II, ISO 27001, ISO 27017 and ISO 42001 certifications, and supports HIPAA-regulated customers."),
        ("Reports", "The SOC 2 Type II report and the penetration test summary are shared with prospects and customers under NDA."),
        ("Audits", "External audits are performed every year by independent auditors."),
    ]),
]

# Questions for the prospect's questionnaire. The last five are deliberately
# NOT covered by the Trust Center documents, so the agent should flag them.
QUESTIONS = [
    ("Data", "In which regions can our data be hosted?"),
    ("Data", "Will data for our EU entities stay within the EU?"),
    ("Data", "How long is our data retained, and what happens to it when the contract ends?"),
    ("Data", "How are we informed about new subprocessors?"),
    ("Encryption", "How is data encrypted in transit?"),
    ("Encryption", "How is data encrypted at rest?"),
    ("Encryption", "How are encryption keys managed and rotated?"),
    ("Encryption", "Do you support customer-managed encryption keys?"),
    ("Access", "How do your employees authenticate to internal systems?"),
    ("Access", "How is access to production systems controlled?"),
    ("Access", "How often are user access rights reviewed?"),
    ("Access", "Can we enforce SSO and automated provisioning for our users?"),
    ("Incidents", "Do you have 24x7 security monitoring?"),
    ("Incidents", "How quickly will you notify us of a data breach?"),
    ("Incidents", "How often is your incident response plan tested?"),
    ("Resilience", "What are your RTO and RPO?"),
    ("Resilience", "How are backups protected?"),
    ("Vulnerabilities", "What are your remediation timelines for critical vulnerabilities?"),
    ("Vulnerabilities", "Do you perform independent penetration tests, and can we see the results?"),
    ("AI", "Is our data used to train AI models?"),
    ("Insurance", "Do you carry cyber liability insurance, and what is the coverage limit?"),
    ("Staffing", "How many full-time employees work in your security team?"),
    ("Physical", "Describe the physical security controls at your own office locations."),
    ("Legal", "Is your source code held in escrow for customers?"),
    ("Payments", "What is your PCI DSS service provider level?"),
]


def make_doc(path, title, review_by, sections):
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run(DISCLAIMER)
    run.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0xA0, 0x40, 0x40)
    doc.add_heading(f"{title} (SAMPLE)", level=1)
    for heading, body in sections:
        doc.add_heading(heading, level=2)
        doc.add_paragraph(body)
    doc.add_paragraph()
    doc.add_paragraph(f"Owner: Security and Compliance. Review by: {review_by}.")
    doc.save(path)


def make_questionnaire(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Questionnaire"
    ws["A1"] = "Contoso Bank: Vendor Security Questionnaire (SAMPLE, fictional)"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "Please answer every question. Attach evidence where available."
    headers = ["ID", "Section", "Question", "Vendor Response", "Comments"]
    ws.append([])
    ws.append(headers)  # header row is row 4
    for c in ws[4]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3A5F")
    for i, (section, q) in enumerate(QUESTIONS, start=1):
        ws.append([f"Q{i:02d}", section, q, "", ""])
    for col, width in zip("ABCDE", [7, 16, 60, 60, 30]):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=5):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path)


def make_nda(path):
    doc = Document()
    doc.add_paragraph(DISCLAIMER).runs[0].italic = True
    doc.add_heading("Mutual Non-Disclosure Agreement (SAMPLE)", level=1)
    doc.add_paragraph("This sample agreement is used only to demonstrate Box Sign in a demo. It has no legal effect.")
    doc.add_paragraph("1. Each party may share confidential security documentation, including audit reports and penetration test summaries, solely to evaluate a potential business relationship.")
    doc.add_paragraph("2. The receiving party will not share confidential information with third parties and will use it only for that evaluation.")
    doc.add_paragraph("3. These obligations last for two years from the date of signature.")
    doc.add_paragraph("\nSigned for the prospect: ______________________   Date: __________")
    doc.save(path)


def make_soc2_placeholder(path):
    doc = Document()
    doc.add_paragraph(DISCLAIMER).runs[0].italic = True
    doc.add_heading("SOC 2 Type II Report (SAMPLE PLACEHOLDER)", level=1)
    doc.add_paragraph("This placeholder stands in for a confidential audit report in the demo. It is shared only after an NDA is signed, through an expiring, view-only link.")
    doc.save(path)


def main():
    (OUT / "01 Trust Center Content").mkdir(parents=True, exist_ok=True)
    (OUT / "Legal").mkdir(exist_ok=True)
    (OUT / "Prospect").mkdir(exist_ok=True)
    for name, title, review_by, sections in TRUST_DOCS:
        make_doc(OUT / "01 Trust Center Content" / name, title, review_by, sections)
    make_soc2_placeholder(OUT / "Legal" / "SOC 2 Type II Report (SAMPLE).docx")
    make_nda(OUT / "Legal" / "Mutual NDA (SAMPLE).docx")
    make_questionnaire(OUT / "Prospect" / "Contoso_Vendor_Security_Questionnaire.xlsx")
    print(f"Wrote demo content to {OUT}")


if __name__ == "__main__":
    main()
