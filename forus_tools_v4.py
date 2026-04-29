"""
forus_tools_v4.py - Forus Toolkit visual tool pages (T1-T4)
All text content is read from the TOOLS sheet at runtime via the `data` dict.
draw_t1/t2/t3/t4 each accept a data dict; defaults are the original hardcoded strings.
build_tools_pdf(selected_ids, data, buf) generates selected tool pages into a BytesIO.
"""

import io
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas as rl_canvas
import math

TEAL   = HexColor('#00424D'); LBLUE = HexColor('#58C5C7'); PINK  = HexColor('#ED1651')
MINT   = HexColor('#5C9C8E'); LIME  = HexColor('#B2C100'); WHITE = HexColor('#FFFFFF')
OFFWHITE = HexColor('#F4F8F8'); MIDGREY = HexColor('#6B7C80')
LTGREY   = HexColor('#E0EAEA'); DKGREY  = HexColor('#2E4A50')
W, H = A4; M = 20*mm; CW = W - 2*M; CX = W / 2


# -- Default text (sourced from TOOLS sheet at runtime) ---

T1_DEFAULTS = {
    "T1_WHY_THIS_MATTERS": (
        "Regulatory compliance is your legal armour. Authorities attacking platforms rarely do so on the real "
        "grounds - they look for technical violations. Completing this check annually means your seat belt is "
        "on before you need it. - Moses Isooba, UNNGOF (Uganda)"),
    "T1_CHECKLIST_LEGAL_STATUS_1": "Registration is current and not due for renewal within 3 months",
    "T1_CHECKLIST_LEGAL_STATUS_2": "All required annual filings (financial, narrative) submitted on time",
    "T1_CHECKLIST_LEGAL_STATUS_3": "Bank account(s) linked to current registered name and address",
    "T1_CHECKLIST_LEGAL_STATUS_4": "Signatory authority on accounts is documented and up to date",
    "T1_CHECKLIST_GOVERNANCE_1":   "Board membership list is current and filed where required",
    "T1_CHECKLIST_GOVERNANCE_2":   "Board meeting minutes recorded and retained for at least 5 years",
    "T1_CHECKLIST_GOVERNANCE_3":   "Conflict-of-interest policy is signed and on file",
    "T1_CHECKLIST_FINANCIAL_1":    "Audited accounts are complete for the most recent financial year",
    "T1_CHECKLIST_FINANCIAL_2":    "Foreign funding is reported as legally required in your jurisdiction",
    "T1_CHECKLIST_FINANCIAL_3":    "Procurement and financial controls policy exists and is followed",
    "T1_CHECKLIST_DIGITAL_1":      "Staff and member data stored in compliance with applicable laws",
    "T1_CHECKLIST_DIGITAL_2":      "Donor and partner contracts are filed and accessible",
    "T1_CHECKLIST_DIGITAL_3":      "A named person holds responsibility for compliance monitoring",
    "T1_CHECKLIST_DUAL_REG_1":     "If operating across borders, secondary registration or banking presence is current",
    "T1_SCORE_1":  "13-14  |  Strong legal posture. Schedule your next review in 12 months.",
    "T1_SCORE_2":  "10-12  |  Good baseline. Address gaps before your next registration renewal cycle.",
    "T1_SCORE_3":  " 7-9   |  Moderate gaps. Prioritise resolution immediately - seek a legal review.",
    "T1_SCORE_4":  " <7    |  Significant vulnerabilities. Contact Forus for urgent legal referral support.",
    "T1_IF_YOU_FIND_GAPS": (
        "Prioritise fixing legal status gaps first - most common basis for regulatory attack. "
        "Contact Forus or your regional coalition for legal referral support."),
}

T2_DEFAULTS = {
    "T2_PROACTIVE_LEGAL_HEALTH": (
        "No immediate crisis. Use the B1 Compliance Self-Check (Tool 1) to identify any gaps before they "
        "become liabilities. Book an annual review with a pro bono legal partner - TrustLaw or ICNL - "
        "before a crisis arises. Prevention is significantly faster than emergency response."),
    "T2_BOX_REGULATORY_BODY": (
        "Contact ICNL (regulatory expertise) or your local bar association pro bono unit first. "
        "Document all correspondence with authorities immediately and in full. Do not respond to official "
        "notices or sign any documents without legal advice. Time matters - document and delay while you find support."),
    "T2_BOX_CRIMINAL_BODY": (
        "Contact Frontline Defenders 24hr emergency line immediately. Then PILnet or TrustLaw for "
        "sustained legal support across the coming weeks. Activate Section 1.4 (Stigmatisation & "
        "Intimidation) in this toolkit immediately. Family financial security is the first priority "
        "in the first 72 hours."),
    "T2_BOX_LITIGATION_BODY": (
        "Contact PILnet for coordination across jurisdictions. Strategic litigation requires a 6-18 month "
        "organisational commitment - assess your leadership capacity and financial stability before "
        "proceeding. See the B3 mechanism directory for full eligibility and regional contact information."),
    "T2_BUYING_TIME": (
        "Request an extension from the authority in writing  ·  Escalate through your regional coalition "
        "or Forus  ·  Document everything  ·  Do not sign anything unilaterally"),
}

T3_DEFAULTS = {
    "T3_STAY_QUIET_LEFT":  (
        "Silence is lower-risk. Notify Forus and your regional coalition privately. "
        "Document the threat. Revisit in 48-72 hours."),
    "T3_STAY_QUIET_RIGHT": (
        "Visibility unlikely to help and carries risk. Escalate privately through Forus "
        "and your donors. Revisit if situation escalates."),
    "T3_PAUSE_SAFETY_FIRST": (
        "Do not publish. Clear any statement with affected staff and partners first. "
        "Seek secure communications guidance (Part 5 protocols) before proceeding."),
    "T3_GO_PUBLIC": (
        "Proceed - but complete the Do-No-Harm Checklist (Tool 4) before publishing. "
        "Frame carefully, protect all sources, and notify Forus in advance."),
}

T4_DEFAULTS = {
    "T4_USE_INSTRUCTION": "Work through every item. A single \"NO\" should trigger a pause and review before publication.",
    # People & Sources
    "T4_PEOPLE_ITEM_1": "Does your statement avoid naming or implying the identity of anyone who has not explicitly consented to being named?",
    "T4_PEOPLE_NOTE_1": "Includes staff, volunteers, beneficiaries, and informal contacts.",
    "T4_PEOPLE_ITEM_2": "Have you removed or anonymised information - locations, dates, organisational names - that could identify individuals?",
    "T4_PEOPLE_NOTE_2": "Even partial details can be dangerous in hostile contexts.",
    "T4_PEOPLE_ITEM_3": "Have all people quoted or referenced reviewed and approved the relevant parts of the statement?",
    "T4_PEOPLE_NOTE_3": "Consent must be explicit, not assumed.",
    # Organisational Safety
    "T4_ORG_ITEM_1": "Does the statement avoid revealing internal details - finances, membership numbers, staff locations - that could be misused?",
    "T4_ORG_NOTE_1": "Authorities and hostile actors mine public statements for operational intelligence.",
    "T4_ORG_ITEM_2": "Have you assessed whether the statement could be used to justify further regulatory action against your organisation?",
    "T4_ORG_NOTE_2": "Consider how it reads to a hostile regulator, not just a sympathetic audience.",
    "T4_ORG_ITEM_3": "Is the platform or channel you are using genuinely secure for this statement?",
    "T4_ORG_NOTE_3": "Email, social media, and WhatsApp carry very different risk profiles. For guidance: Access Now Digital Security Helpline - accessnow.org/help",
    # Partners & Network
    "T4_PARTNERS_ITEM_1": "Have you notified any partner organisations mentioned or implicated in the statement before it is published?",
    "T4_PARTNERS_NOTE_1": "Surprise exposure can damage trust and create secondary risk for others.",
    "T4_PARTNERS_ITEM_2": "Have you considered whether publication could create legal or regulatory risk for any Forus member or partner organisation?",
    # Timing & Framing
    "T4_TIMING_ITEM_1": "Is this the right moment? Have you checked whether publishing now would disrupt any ongoing legal process, negotiation, or dialogue?",
    "T4_TIMING_NOTE_1": "A statement correct in content can be seriously harmful in timing.",
    "T4_TIMING_ITEM_2": "Does the statement avoid language that could characterise your organisation as hostile, foreign-funded, or a national security risk?",
    "T4_TIMING_NOTE_2": "Framing matters as much as the facts themselves.",
    "T4_TIMING_ITEM_3": "Has the statement been reviewed by at least one person outside your immediate team?",
    "T4_TIMING_NOTE_3": "A second pair of eyes catches risks that are invisible from inside the situation.",
    "T4_IF_ANY_NO": (
        "Pause publication. Address the specific risk identified before proceeding. "
        "If unsure, consult your Forus focal point or regional coalition communications lead."),
}
