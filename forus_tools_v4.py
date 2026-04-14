"""
forus_tools_v4.py — Forus Toolkit visual tool pages (T1–T4)
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


# ── Default text (sourced from TOOLS sheet at runtime) ───────────────────────

T1_DEFAULTS = {
    "T1_WHY_THIS_MATTERS": (
        "Regulatory compliance is your legal armour. Authorities attacking platforms rarely do so on the real "
        "grounds — they look for technical violations. Completing this check annually means your seat belt is "
        "on before you need it. (Moses, Uganda)"),
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
    "T1_SCORE_1":  "13–14  ✓  |  Strong legal posture. Schedule your next review in 12 months.",
    "T1_SCORE_2":  "10–12  ✓  |  Good baseline. Address gaps before your next registration renewal cycle.",
    "T1_SCORE_3":  " 7–9   ✓  |  Moderate gaps. Prioritise resolution immediately — seek a legal review.",
    "T1_SCORE_4":  " <7    ✓  |  Significant vulnerabilities. Contact Forus for urgent legal referral support.",
    "T1_IF_YOU_FIND_GAPS": (
        "Prioritise fixing legal status gaps first — most common basis for regulatory attack. "
        "Contact Forus or your regional coalition for legal referral support."),
}

T2_DEFAULTS = {
    "T2_PROACTIVE_LEGAL_HEALTH": (
        "No immediate crisis. Use the B1 Compliance Self-Check (Tool 1) to identify any gaps before they "
        "become liabilities. Book an annual review with a pro bono legal partner — TrustLaw or ICNL — "
        "before a crisis arises. Prevention is significantly faster than emergency response."),
    "T2_BOX_REGULATORY_BODY": (
        "Contact ICNL (regulatory expertise) or your local bar association pro bono unit first. "
        "Document all correspondence with authorities immediately and in full. Do not respond to official "
        "notices or sign any documents without legal advice. Time matters — document and delay while you find support."),
    "T2_BOX_CRIMINAL_BODY": (
        "Contact Frontline Defenders 24hr emergency line immediately. Then PILnet or TrustLaw for "
        "sustained legal support across the coming weeks. Activate Scenario 1.5 (Detention) in this "
        "toolkit immediately. Family financial security is the first priority in the first 72 hours."),
    "T2_BOX_LITIGATION_BODY": (
        "Contact PILnet for coordination across jurisdictions. Strategic litigation requires a 6–18 month "
        "organisational commitment — assess your leadership capacity and financial stability before "
        "proceeding. See the B3 mechanism directory for full eligibility and regional contact information."),
    "T2_BUYING_TIME": (
        "Request an extension from the authority in writing  ·  Escalate through your regional coalition "
        "or Forus  ·  Document everything  ·  Do not sign anything unilaterally"),
}

T3_DEFAULTS = {
    "T3_STAY_QUIET_LEFT":  (
        "Silence is lower-risk. Notify Forus and your regional coalition privately. "
        "Document the threat. Revisit in 48–72 hours."),
    "T3_STAY_QUIET_RIGHT": (
        "Visibility unlikely to help and carries risk. Escalate privately through Forus "
        "and your donors. Revisit if situation escalates."),
    "T3_PAUSE_SAFETY_FIRST": (
        "Do not publish. Clear any statement with affected staff and partners first. "
        "Seek secure communications guidance (Part 5 protocols) before proceeding."),
    "T3_GO_PUBLIC": (
        "Proceed — but complete the Do-No-Harm Checklist (Tool 4) before publishing. "
        "Frame carefully, protect all sources, and notify Forus in advance."),
}

T4_DEFAULTS = {
    "T4_USE_INSTRUCTION": "Work through every item. A single \"NO\" should trigger a pause and review before publication.",
    # People & Sources
    "T4_PEOPLE_ITEM_1": "Does your statement avoid naming or implying the identity of anyone who has not explicitly consented to being named?",
    "T4_PEOPLE_NOTE_1": "Includes staff, volunteers, beneficiaries, and informal contacts.",
    "T4_PEOPLE_ITEM_2": "Have you removed or anonymised information — locations, dates, organisational names — that could identify individuals?",
    "T4_PEOPLE_NOTE_2": "Even partial details can be dangerous in hostile contexts.",
    "T4_PEOPLE_ITEM_3": "Have all people quoted or referenced reviewed and approved the relevant parts of the statement?",
    "T4_PEOPLE_NOTE_3": "Consent must be explicit, not assumed.",
    # Organisational Safety
    "T4_ORG_ITEM_1": "Does the statement avoid revealing internal details — finances, membership numbers, staff locations — that could be misused?",
    "T4_ORG_NOTE_1": "Authorities and hostile actors mine public statements for operational intelligence.",
    "T4_ORG_ITEM_2": "Have you assessed whether the statement could be used to justify further regulatory action against your organisation?",
    "T4_ORG_NOTE_2": "Consider how it reads to a hostile regulator, not just a sympathetic audience.",
    "T4_ORG_ITEM_3": "Is the platform or channel you are using genuinely secure for this statement?",
    "T4_ORG_NOTE_3": "Email, social media, and WhatsApp carry very different risk profiles. For guidance: Access Now Digital Security Helpline — accessnow.org/help",
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


# ── Localised UI strings (structural labels, headings, button text) ───────────

_UI = {
    "EN": {
        "tool_tag":        "TOOL {num}  ·  {sec}",
        "toolkit_name":    "FORUS RESILIENCE TOOLKIT",
        "footer_url":      "forus-international.org",
        "why_matters":     "WHY THIS MATTERS",
        "how_to_use_t1":   "HOW TO USE:  Tick each item your organisation can confirm today. Score one point per tick. Flag any gaps for immediate action.",
        "scoring_guide":   "SCORING GUIDE",
        "if_gaps":         "IF YOU FIND GAPS",
        "legal_status":    "LEGAL STATUS",
        "governance":      "GOVERNANCE",
        "financial":       "FINANCIAL",
        "digital_ops":     "DIGITAL & OPERATIONAL",
        "dual_reg":        "DUAL REGISTRATION (WHERE RELEVANT)",
        "t1_footer":       "Tool 1 of 4  ·  Complete annually — or whenever your legal environment changes significantly",
        "need_legal":      "You need legal support",
        "imm_threat":      "Immediate threat or legal deadline?",
        "kind_threat":     "What kind of threat?",
        "reg_attack":      "REGULATORY ATTACK",
        "crim_det":        "CRIMINAL / DETENTION",
        "strat_lit":       "STRATEGIC LITIGATION",
        "comp_def":        "COMPLIANCE DEFENCE",
        "crim_def":        "CRIMINAL DEFENCE",
        "proactive_legal": "PROACTIVE LEGAL HEALTH",
        "buying_time_lbl": "BUYING TIME when legal support is not immediately available",
        "t2_footer":       "Tool 2 of 4  ·  See Part 3 mechanism directory for full eligibility details and contact information",
        "going_public":    "You are considering going public",
        "pub_domain":      "Is the threat already in the public domain?",
        "silence_complic": "Would silence look like\ncomplicity or harm credibility?",
        "visibility_prot": "Would visibility bring\nmeaningful pressure or protection?",
        "endanger":        "Could speaking out directly endanger\nstaff, members, or partners?",
        "yes":             "YES",
        "no_lbl":          "NO",
        "stay_quiet":      "STAY QUIET (FOR NOW)",
        "pause":           "PAUSE —",
        "safety_first":    "SAFETY FIRST",
        "go_public":       "GO PUBLIC",
        "t3_footer":       "Tool 3 of 4  ·  Revisit this tree if your context changes — situations evolve quickly",
        "use_after_t3":    "USE AFTER TOOL 3 — IF YOU HAVE DECIDED TO GO PUBLIC",
        "people_sources":  "PEOPLE & SOURCES",
        "org_safety":      "ORGANISATIONAL SAFETY",
        "partners_net":    "PARTNERS & NETWORK",
        "timing_framing":  "TIMING & FRAMING",
        "if_any_no":       "IF ANY ANSWER IS NO",
        "t4_footer":       "Tool 4 of 4  ·  Keep a completed copy of this checklist on file alongside each statement published",
    },
    "FR": {
        "tool_tag":        "OUTIL {num}  ·  {sec}",
        "toolkit_name":    "BOITE A OUTILS RESILIENCE FORUS",
        "footer_url":      "forus-international.org",
        "why_matters":     "POURQUOI C'EST IMPORTANT",
        "how_to_use_t1":   "MODE D'EMPLOI :  Cochez chaque element que votre organisation peut confirmer. Un point par coche. Signalez les lacunes pour action immediate.",
        "scoring_guide":   "GUIDE DE NOTATION",
        "if_gaps":         "SI DES LACUNES SONT IDENTIFIEES",
        "legal_status":    "STATUT JURIDIQUE",
        "governance":      "GOUVERNANCE",
        "financial":       "FINANCES",
        "digital_ops":     "NUMERIQUE & OPERATIONNEL",
        "dual_reg":        "DOUBLE ENREGISTREMENT (LE CAS ECHEANT)",
        "t1_footer":       "Outil 1 sur 4  ·  A completer annuellement ou lorsque votre environnement juridique evolue",
        "need_legal":      "Vous avez besoin d'un soutien juridique",
        "imm_threat":      "Menace immediate ou delai juridique ?",
        "kind_threat":     "Quel type de menace ?",
        "reg_attack":      "ATTAQUE REGLEMENTAIRE",
        "crim_det":        "PENAL / DETENTION",
        "strat_lit":       "CONTENTIEUX STRATEGIQUE",
        "comp_def":        "DEFENSE DE CONFORMITE",
        "crim_def":        "DEFENSE PENALE",
        "proactive_legal": "SANTE JURIDIQUE PREVENTIVE",
        "buying_time_lbl": "GAGNER DU TEMPS en attendant un soutien juridique",
        "t2_footer":       "Outil 2 sur 4  ·  Voir le repertoire des mecanismes (Partie 3) pour les criteres d'eligibilite",
        "going_public":    "Vous envisagez de rendre public",
        "pub_domain":      "La menace est-elle deja dans le domaine public ?",
        "silence_complic": "Le silence semblerait-il\nune complicite ou nuirait a la credibilite ?",
        "visibility_prot": "La visibilite apporterait-elle\nune pression ou protection significative ?",
        "endanger":        "Prendre la parole mettrait-il\nen danger le personnel ou membres ?",
        "yes":             "OUI",
        "no_lbl":          "NON",
        "stay_quiet":      "RESTER SILENCIEUX (POUR L'INSTANT)",
        "pause":           "PAUSE -",
        "safety_first":    "LA SECURITE AVANT TOUT",
        "go_public":       "RENDRE PUBLIC",
        "t3_footer":       "Outil 3 sur 4  ·  Revenez a cet arbre si le contexte evolue",
        "use_after_t3":    "A UTILISER APRES L'OUTIL 3 - SI VOUS AVEZ DECIDE DE RENDRE PUBLIC",
        "people_sources":  "PERSONNES ET SOURCES",
        "org_safety":      "SECURITE ORGANISATIONNELLE",
        "partners_net":    "PARTENAIRES ET RESEAU",
        "timing_framing":  "CALENDRIER ET CADRAGE",
        "if_any_no":       "SI UNE REPONSE EST NON",
        "t4_footer":       "Outil 4 sur 4  ·  Conservez une copie complete avec chaque declaration publiee",
    },
    "ES": {
        "tool_tag":        "HERRAMIENTA {num}  ·  {sec}",
        "toolkit_name":    "KIT DE HERRAMIENTAS FORUS",
        "footer_url":      "forus-international.org",
        "why_matters":     "POR QUE ES IMPORTANTE",
        "how_to_use_t1":   "COMO USAR:  Marque cada elemento que su organizacion pueda confirmar hoy. Un punto por marca. Senale las brechas para accion inmediata.",
        "scoring_guide":   "GUIA DE PUNTUACION",
        "if_gaps":         "SI ENCUENTRA BRECHAS",
        "legal_status":    "ESTADO JURIDICO",
        "governance":      "GOBERNANZA",
        "financial":       "FINANZAS",
        "digital_ops":     "DIGITAL Y OPERACIONAL",
        "dual_reg":        "DOBLE REGISTRO (CUANDO SEA PERTINENTE)",
        "t1_footer":       "Herramienta 1 de 4  ·  Completar anualmente o cuando su entorno juridico cambie",
        "need_legal":      "Necesita apoyo juridico",
        "imm_threat":      "Amenaza inmediata o plazo juridico?",
        "kind_threat":     "Que tipo de amenaza?",
        "reg_attack":      "ATAQUE REGULATORIO",
        "crim_det":        "PENAL / DETENCION",
        "strat_lit":       "LITIGIO ESTRATEGICO",
        "comp_def":        "DEFENSA DE CUMPLIMIENTO",
        "crim_def":        "DEFENSA PENAL",
        "proactive_legal": "SALUD JURIDICA PROACTIVA",
        "buying_time_lbl": "GANAR TIEMPO cuando el apoyo juridico no esta disponible aun",
        "t2_footer":       "Herramienta 2 de 4  ·  Ver directorio de mecanismos (Parte 3) para elegibilidad y contactos",
        "going_public":    "Esta considerando hacerlo publico",
        "pub_domain":      "La amenaza ya es de dominio publico?",
        "silence_complic": "El silencio pareceria complicidad\no danaria la credibilidad?",
        "visibility_prot": "La visibilidad aportaria\npresion o proteccion significativa?",
        "endanger":        "Hablar publicamente pondria en peligro\nal personal, miembros o socios?",
        "yes":             "SI",
        "no_lbl":          "NO",
        "stay_quiet":      "GUARDAR SILENCIO (POR AHORA)",
        "pause":           "PAUSA -",
        "safety_first":    "LA SEGURIDAD PRIMERO",
        "go_public":       "HACER PUBLICO",
        "t3_footer":       "Herramienta 3 de 4  ·  Revise este arbol si el contexto cambia",
        "use_after_t3":    "USAR DESPUES DE LA HERRAMIENTA 3 - SI HA DECIDIDO HACER PUBLICO",
        "people_sources":  "PERSONAS Y FUENTES",
        "org_safety":      "SEGURIDAD ORGANIZACIONAL",
        "partners_net":    "SOCIOS Y RED",
        "timing_framing":  "MOMENTO Y ENFOQUE",
        "if_any_no":       "SI ALGUNA RESPUESTA ES NO",
        "t4_footer":       "Herramienta 4 de 4  ·  Archive una copia completa con cada declaracion publicada",
    },
}


# ── Drawing utilities ─────────────────────────────────────────────────────────

def _hdr(c, num, sec, title, sub=None, _u=None):
    if _u is None: _u = _UI["EN"]
    bh = 48*mm; c.setFillColor(TEAL); c.rect(0, H-bh, W, bh, fill=1, stroke=0)
    tag = _u["tool_tag"].format(num=num, sec=sec)
    tw = c.stringWidth(tag, 'Helvetica-Bold', 7) + 6*mm
    c.setFillColor(LBLUE); c.roundRect(M, H-11*mm, tw, 7*mm, 1*mm, fill=1, stroke=0)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 7); c.drawString(M+3*mm, H-7.5*mm, tag)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 21); c.drawString(M, H-27*mm, title)
    if sub: c.setFillColor(LBLUE); c.setFont('Helvetica', 9); c.drawString(M, H-35*mm, sub)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 9)
    c.drawRightString(W-M, H-7.5*mm, _u["toolkit_name"])


def _ftr(c, note, _u=None):
    if _u is None: _u = _UI["EN"]
    c.setFillColor(TEAL); c.rect(0, 0, W, 9*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica', 7)
    c.drawString(M, 3*mm, note); c.drawRightString(W-M, 3*mm, _u["footer_url"])


def _dp(c, pts, fc=None, sc=None, lw=1):
    if sc: c.setStrokeColor(sc); c.setLineWidth(lw)
    if fc: c.setFillColor(fc)
    p = c.beginPath(); p.moveTo(*pts[0])
    for pt in pts[1:]: p.lineTo(*pt)
    p.close(); c.drawPath(p, fill=1 if fc else 0, stroke=1 if sc else 0)


def _av(c, x, y0, y1, col=LBLUE, lw=1.2):
    c.setStrokeColor(col); c.setLineWidth(lw); c.line(x, y0, x, y1)
    _dp(c, [(x, y1), (x-2*mm, y1+4*mm), (x+2*mm, y1+4*mm)], fc=col)


def _ah(c, x0, x1, y, col=LBLUE, lw=1.2):
    c.setStrokeColor(col); c.setLineWidth(lw); c.line(x0, y, x1, y)
    if x1 > x0: _dp(c, [(x1, y), (x1-4*mm, y-2*mm), (x1-4*mm, y+2*mm)], fc=col)
    else:        _dp(c, [(x1, y), (x1+4*mm, y-2*mm), (x1+4*mm, y+2*mm)], fc=col)


def _wl(c, txt, fn, fs, mw):
    wds = txt.split(); out = []; ln = ''
    for w2 in wds:
        t = (ln + ' ' + w2).strip()
        if c.stringWidth(t, fn, fs) <= mw: ln = t
        else:
            if ln: out.append(ln)
            ln = w2
    if ln: out.append(ln)
    return out


def _tb(c, txt, x, y, mw, fn, fs, col=DKGREY, ld=None):
    if ld is None: ld = fs * 1.3
    c.setFillColor(col); c.setFont(fn, fs)
    for ln in _wl(c, txt, fn, fs, mw):
        c.drawString(x, y, ln); y -= ld
    return y


def _lbl(c, x, y, txt, col=MIDGREY, fs=7, align='l'):
    c.setFillColor(col); c.setFont('Helvetica-BoldOblique', fs)
    {'l': c.drawString, 'r': c.drawRightString, 'c': c.drawCentredString}[align](x, y, txt)


def _gh(c, x, y, w, title, col=TEAL):
    c.setFillColor(col); c.roundRect(x, y-5.5*mm, w, 5.5*mm, 1*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 7); c.drawString(x+3*mm, y-3.8*mm, title)


def _cbx(c, x, y, sz=3.8*mm, col=TEAL):
    c.setStrokeColor(col); c.setLineWidth(0.8); c.rect(x, y, sz, sz, fill=0, stroke=1)


def _rrb(c, x, y, w, h, r, bg, lines=None, fg=WHITE, fs=8.5, bold=True):
    c.setFillColor(bg); c.roundRect(x, y, w, h, r, fill=1, stroke=0)
    if lines:
        fn = 'Helvetica-Bold' if bold else 'Helvetica'
        c.setFillColor(fg); lh = fs * 1.35; n = len(lines)
        ty = y + h/2 + (n-1)*lh/2
        for ln in lines: c.setFont(fn, fs); c.drawCentredString(x+w/2, ty, ln); ty -= lh


def _dmnd(c, cx, cy, hw, hh, fc=TEAL):
    _dp(c, [(cx, cy+hh), (cx+hw, cy), (cx, cy-hh), (cx-hw, cy)], fc=fc)


# ── Tool 1 ───────────────────────────────────────────────────────────────────

def draw_t1(c, data=None, language="EN"):
    d = dict(T1_DEFAULTS)
    if data: d.update(data)
    _u = _UI.get(language.upper(), _UI["EN"])

    _hdr(c, '1 of 4', 'B1 · CRISIS SCENARIOS', 'Compliance Readiness Self-Check',
         'Complete before a crisis — not during one', _u=_u)
    ty = H - 48*mm - 6*mm

    ib_h = 22*mm
    c.setFillColor(OFFWHITE); c.roundRect(M, ty-ib_h, CW, ib_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 8.5)
    c.drawString(M+4*mm, ty-6*mm, _u["why_matters"])
    _tb(c, d["T1_WHY_THIS_MATTERS"], M+4*mm, ty-13*mm, CW-8*mm, 'Helvetica', 8, col=MIDGREY, ld=10.5)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 7.5)
    c.drawString(M, ty-ib_h-5*mm, _u["how_to_use_t1"])

    cw = (CW - 6*mm) / 2; GAP = 7

    def draw_col(c, x, sy, groups):
        y = sy
        for gt, items in groups:
            _gh(c, x, y, cw, gt); y -= 8*mm
            for txt in items:
                sz = 3.8*mm; _cbx(c, x+1*mm, y-sz-0.5*mm, sz)
                lns = _wl(c, txt, 'Helvetica', 8, cw-8*mm)
                for ln in lns:
                    c.setFillColor(DKGREY); c.setFont('Helvetica', 8)
                    c.drawString(x+7*mm, y-2*mm, ln); y -= 10
                y -= GAP
            y -= 4*mm
        return y

    g1 = [
        (_u["legal_status"], [
            d["T1_CHECKLIST_LEGAL_STATUS_1"], d["T1_CHECKLIST_LEGAL_STATUS_2"],
            d["T1_CHECKLIST_LEGAL_STATUS_3"], d["T1_CHECKLIST_LEGAL_STATUS_4"],
        ]),
        (_u["governance"], [
            d["T1_CHECKLIST_GOVERNANCE_1"], d["T1_CHECKLIST_GOVERNANCE_2"],
            d["T1_CHECKLIST_GOVERNANCE_3"],
        ]),
    ]
    g2 = [
        (_u["financial"], [
            d["T1_CHECKLIST_FINANCIAL_1"], d["T1_CHECKLIST_FINANCIAL_2"],
            d["T1_CHECKLIST_FINANCIAL_3"],
        ]),
        (_u["digital_ops"], [
            d["T1_CHECKLIST_DIGITAL_1"], d["T1_CHECKLIST_DIGITAL_2"],
            d["T1_CHECKLIST_DIGITAL_3"],
        ]),
        (_u["dual_reg"], [d["T1_CHECKLIST_DUAL_REG_1"]]),
    ]

    cs = ty - ib_h - 10*mm
    bot1 = draw_col(c, M, cs, g1)
    bot2 = draw_col(c, M+cw+6*mm, cs, g2)
    content_bot = min(bot1, bot2)

    # Scoring guide
    sg_h = 22*mm; sg_y = content_bot - 8*mm - sg_h
    c.setFillColor(LTGREY); c.roundRect(M, sg_y, CW, sg_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 7.5)
    c.drawString(M+4*mm, sg_y+sg_h-5*mm, _u["scoring_guide"])

    def _parse_score(raw):
        parts = raw.split('|', 1)
        return (parts[0].strip(), parts[1].strip() if len(parts) > 1 else '')

    scores = [_parse_score(d[f"T1_SCORE_{i}"]) for i in range(1, 5)]
    sy2 = sg_y + sg_h - 12*mm
    for score, desc in scores:
        c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 7.5); c.drawString(M+4*mm, sy2, score)
        c.setFillColor(DKGREY); c.setFont('Helvetica', 7.5); c.drawString(M+30*mm, sy2, desc)
        sy2 -= 9.5

    # IF YOU FIND GAPS
    ah2 = 22*mm; ay = sg_y - 6*mm - ah2
    c.setFillColor(PINK); c.roundRect(M, ay, CW, ah2, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 8.5)
    c.drawString(M+4*mm, ay+16*mm, _u["if_gaps"])
    _tb(c, d["T1_IF_YOU_FIND_GAPS"], M+4*mm, ay+10*mm, CW-8*mm, 'Helvetica', 8, col=WHITE, ld=10)
    _ftr(c, _u["t1_footer"], _u=_u)


# ── Tool 2 ───────────────────────────────────────────────────────────────────

def draw_t2(c, data=None, language="EN"):
    d = dict(T2_DEFAULTS)
    if data: d.update(data)
    _u = _UI.get(language.upper(), _UI["EN"])

    _hdr(c, '2 of 4', 'B3 · LEGAL SUPPORT', 'Legal Support Decision Tree',
         'What kind of legal support do you need?', _u=_u)
    ty = H - 48*mm - 10*mm

    # ── Compute box geometry first so yline_x can align with box 1 centre ──────
    bw = (CW - 6*mm) / 3; bh = 65*mm
    bxs = [M, M + bw + 3*mm, M + 2*(bw + 3*mm)]
    # yline_x = centre of left outcome box; q2hw capped so diamond stays within margin
    yline_x = bxs[0] + bw / 2                        # ≈ 47 mm
    q2hw    = min(26*mm, yline_x - M - 2*mm)         # ≈ 25 mm — fits within left margin

    q1cx = CX - 15*mm; q1w = 78*mm; q1h = 13*mm
    sw = 85*mm; sh = 12*mm
    _rrb(c, CX-sw/2, ty-sh, sw, sh, 5*mm, TEAL, [_u["need_legal"]], fs=9.5)
    _av(c, CX, ty-sh, ty-sh-13*mm)

    q1y = ty - sh - 13*mm - q1h
    _rrb(c, q1cx-q1w/2, q1y, q1w, q1h, 2*mm, DKGREY, [_u["imm_threat"]], fs=8.5)

    # YES: short horizontal run from Q1 left edge → yline_x, then down to Q2 diamond
    c.setStrokeColor(PINK); c.setLineWidth(1.2)
    c.line(q1cx-q1w/2, q1y+q1h/2, yline_x, q1y+q1h/2)
    _lbl(c, q1cx-q1w/2-2*mm, q1y+q1h/2+2*mm, _u["yes"], col=PINK, fs=7, align='r')

    pro_x = q1cx + q1w/2 + 8*mm; pro_w = W - M - pro_x; pro_h = 55*mm
    pro_y = q1y + q1h - pro_h
    _ah(c, q1cx+q1w/2, pro_x, q1y+q1h*0.6, col=MINT)
    _lbl(c, q1cx+q1w/2+2*mm, q1y+q1h*0.6+2*mm, _u["no_lbl"], col=MINT, fs=7)
    c.setFillColor(MINT); c.roundRect(pro_x, pro_y, pro_w, pro_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 8)
    c.drawString(pro_x+3*mm, pro_y+pro_h-6*mm, _u["proactive_legal"])
    _tb(c, d["T2_PROACTIVE_LEGAL_HEALTH"], pro_x+3*mm, pro_y+pro_h-14*mm,
        pro_w-6*mm, 'Helvetica', 7.8, col=WHITE, ld=10.5)

    q2hh = 13*mm
    q2y = q1y - 2*mm; q2cy = q2y - 16*mm
    _av(c, yline_x, q1y+q1h/2, q2cy+q2hh, col=PINK)
    _dmnd(c, yline_x, q2cy, q2hw, q2hh, fc=PINK)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 7.5)
    c.drawCentredString(yline_x, q2cy+1.5*mm, _u["kind_threat"])

    branch_y = q2cy - q2hh - 12*mm
    box_top  = branch_y - 6*mm
    c.setStrokeColor(PINK); c.setLineWidth(1)
    c.line(bxs[0]+bw/2, branch_y, bxs[2]+bw/2, branch_y)
    c.line(yline_x, q2cy-q2hh, yline_x, branch_y)

    boxes = [
        (bxs[0], _u["reg_attack"], TEAL, _u["comp_def"], d["T2_BOX_REGULATORY_BODY"]),
        (bxs[1], _u["crim_det"],   PINK, _u["crim_def"], d["T2_BOX_CRIMINAL_BODY"]),
        (bxs[2], _u["strat_lit"],  MINT, _u["strat_lit"], d["T2_BOX_LITIGATION_BODY"]),
    ]
    for bxl, btitle, bcol, bhead, btext in boxes:
        bcx = bxl + bw/2
        _av(c, bcx, branch_y, box_top, col=PINK)
        c.setFillColor(PINK); c.setFont('Helvetica-Bold', 6.5)
        c.drawCentredString(bcx, branch_y+2*mm, btitle)
        c.setFillColor(bcol); c.roundRect(bxl, box_top-bh, bw, bh, 2*mm, fill=1, stroke=0)
        c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 8)
        c.drawCentredString(bxl+bw/2, box_top-bh+bh-6*mm, bhead)
        _tb(c, btext, bxl+3*mm, box_top-bh+bh-14*mm, bw-6*mm, 'Helvetica', 7.5, col=WHITE, ld=10)

    outcome_bottom = box_top - bh; bth = 16*mm; by = outcome_bottom - 12*mm - bth
    c.setFillColor(OFFWHITE); c.roundRect(M, by, CW, bth, 2*mm, fill=1, stroke=0)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 8)
    c.drawString(M+4*mm, by+10*mm, _u["buying_time_lbl"])
    _tb(c, d["T2_BUYING_TIME"], M+4*mm, by+3.5*mm, CW-8*mm, 'Helvetica', 7.8, col=MIDGREY, ld=10)
    _ftr(c, _u["t2_footer"], _u=_u)


# ── Tool 3 ───────────────────────────────────────────────────────────────────

def draw_t3(c, data=None, language="EN"):
    d = dict(T3_DEFAULTS)
    if data: d.update(data)
    _u = _UI.get(language.upper(), _UI["EN"])

    _hdr(c, '3 of 4', 'B5 · SAFE ADVOCACY & COMMUNICATIONS', '"Go public, or stay quiet?"',
         'Use this decision tree before making a public statement in a sensitive context', _u=_u)
    ty = H - 48*mm - 8*mm

    NW = 90*mm; NH = 12*mm; BW = 62*mm; BH = 14*mm
    Q4W = 82*mm; Q4H = 12*mm; SQW = 52*mm; SQH = 38*mm; OW = 46*mm; OH = 40*mm
    q3x = CX - 42*mm; q2x = CX + 42*mm

    y0 = ty; y1 = y0 - NH - 14*mm; br = y1 - NH - 16*mm
    y2 = br - 8*mm - BH; y2t = y2 + BH; y2cy = y2 + BH/2
    sq_top = y2 - 10*mm; sq_bot = sq_top - SQH
    q4_top = sq_bot - 16*mm; q4_bot = q4_top - Q4H; q4_cy = q4_bot + Q4H/2
    out_top = q4_bot - 10*mm; out_bot = out_top - OH

    _rrb(c, CX-NW/2, y0-NH, NW, NH, 5*mm, TEAL, [_u["going_public"]], fs=9)
    _av(c, CX, y0-NH, y1)
    _rrb(c, CX-NW/2, y1-NH, NW, NH, 2*mm, DKGREY, [_u["pub_domain"]], fs=8.5)

    c.setStrokeColor(LBLUE); c.setLineWidth(1.2)
    c.line(CX, y1-NH, CX, br); c.line(q3x, br, q2x, br)
    _av(c, q3x, br, y2t, col=LBLUE); _av(c, q2x, br, y2t, col=LBLUE)
    _lbl(c, q3x, br+2*mm, _u["no_lbl"], align='c'); _lbl(c, q2x, br+2*mm, _u["yes"], align='c')

    _rrb(c, q3x-BW/2, y2, BW, BH, 2*mm, DKGREY, _u["silence_complic"].split('\n'), fs=7.5)
    _rrb(c, q2x-BW/2, y2, BW, BH, 2*mm, DKGREY, _u["visibility_prot"].split('\n'), fs=7.5)

    c.setStrokeColor(MINT); c.setLineWidth(1.2)
    c.line(q3x+BW/2, y2cy, CX, y2cy); c.line(q2x-BW/2, y2cy, CX, y2cy)
    _lbl(c, CX, y2cy+3*mm, _u["yes"], col=MINT, align='c')
    _av(c, CX, y2cy, q4_top, col=MINT)
    _av(c, q3x, y2, sq_top, col=PINK); _lbl(c, q3x-12*mm, y2-5*mm, _u["no_lbl"], col=PINK)
    _av(c, q2x, y2, sq_top, col=PINK); _lbl(c, q2x+2*mm, y2-5*mm, _u["no_lbl"], col=PINK)

    for sqx, sqtxt in [
        (q3x-SQW/2, d["T3_STAY_QUIET_LEFT"]),
        (q2x-SQW/2, d["T3_STAY_QUIET_RIGHT"]),
    ]:
        c.setFillColor(MINT); c.roundRect(sqx, sq_bot, SQW, SQH, 2*mm, fill=1, stroke=0)
        c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 8.5)
        c.drawCentredString(sqx+SQW/2, sq_bot+SQH-6*mm, _u["stay_quiet"])
        _tb(c, sqtxt, sqx+3*mm, sq_bot+SQH-13*mm, SQW-6*mm, 'Helvetica', 7.5, col=WHITE, ld=9.5)

    _rrb(c, CX-Q4W/2, q4_bot, Q4W, Q4H, 2*mm, DKGREY, _u["endanger"].split('\n'), fs=8)

    _av(c, CX-Q4W/3, q4_bot, out_top, col=PINK)
    _lbl(c, CX-Q4W/3-12*mm, q4_bot-5*mm, _u["yes"], col=PINK)
    _av(c, CX+Q4W/3, q4_bot, out_top, col=TEAL)
    _lbl(c, CX+Q4W/3+2*mm, q4_bot-5*mm, _u["no_lbl"], col=TEAL)

    px = M
    c.setFillColor(PINK); c.roundRect(px, out_bot, OW, OH, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 9)
    c.drawCentredString(px+OW/2, out_bot+OH-6*mm, _u["pause"])
    c.drawCentredString(px+OW/2, out_bot+OH-14*mm, _u["safety_first"])
    _tb(c, d["T3_PAUSE_SAFETY_FIRST"], px+3*mm, out_bot+OH-21*mm, OW-6*mm, 'Helvetica', 7.5, col=WHITE, ld=9.5)

    gx = W - M - OW
    c.setFillColor(TEAL); c.roundRect(gx, out_bot, OW, OH, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 9)
    c.drawCentredString(gx+OW/2, out_bot+OH-6*mm, _u["go_public"])
    _tb(c, d["T3_GO_PUBLIC"], gx+3*mm, out_bot+OH-14*mm, OW-6*mm, 'Helvetica', 7.5, col=WHITE, ld=9.5)
    _ftr(c, _u["t3_footer"], _u=_u)


# ── Tool 4 ───────────────────────────────────────────────────────────────────

def draw_t4(c, data=None, language="EN"):
    d = dict(T4_DEFAULTS)
    if data: d.update(data)
    _u = _UI.get(language.upper(), _UI["EN"])

    _hdr(c, '4 of 4', 'B5 · SAFE ADVOCACY & COMMUNICATIONS', 'Do-No-Harm Communications Checklist',
         'Complete this before publishing any statement in a sensitive or hostile environment', _u=_u)
    ty = H - 48*mm - 6*mm

    ib_h = 18*mm
    c.setFillColor(OFFWHITE); c.roundRect(M, ty-ib_h, CW, ib_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 8.5)
    c.drawString(M+4*mm, ty-6*mm, _u["use_after_t3"])
    c.setFillColor(MIDGREY); c.setFont('Helvetica', 8.5)
    c.drawString(M+4*mm, ty-13*mm, d["T4_USE_INSTRUCTION"])

    grps = [
        (_u["people_sources"], PINK, [
            (d["T4_PEOPLE_ITEM_1"], d["T4_PEOPLE_NOTE_1"]),
            (d["T4_PEOPLE_ITEM_2"], d["T4_PEOPLE_NOTE_2"]),
            (d["T4_PEOPLE_ITEM_3"], d["T4_PEOPLE_NOTE_3"]),
        ]),
        (_u["org_safety"], TEAL, [
            (d["T4_ORG_ITEM_1"], d["T4_ORG_NOTE_1"]),
            (d["T4_ORG_ITEM_2"], d["T4_ORG_NOTE_2"]),
            (d["T4_ORG_ITEM_3"], d["T4_ORG_NOTE_3"]),
        ]),
        (_u["partners_net"], LBLUE, [
            (d["T4_PARTNERS_ITEM_1"], d["T4_PARTNERS_NOTE_1"]),
            (d["T4_PARTNERS_ITEM_2"], ''),
        ]),
        (_u["timing_framing"], MINT, [
            (d["T4_TIMING_ITEM_1"], d["T4_TIMING_NOTE_1"]),
            (d["T4_TIMING_ITEM_2"], d["T4_TIMING_NOTE_2"]),
            (d["T4_TIMING_ITEM_3"], d["T4_TIMING_NOTE_3"]),
        ]),
    ]

    cw2 = (CW - 5*mm) / 2; GAP = 5

    def draw_grp(c, x, y, grp):
        title, col, items = grp
        _gh(c, x, y, cw2, title, col=col); y -= 8*mm
        for main, sub in items:
            sz = 3.8*mm; _cbx(c, x+1*mm, y-sz-0.5*mm, sz, col=col)
            for ln in _wl(c, main, 'Helvetica-Bold', 7.8, cw2-8*mm):
                c.setFillColor(DKGREY); c.setFont('Helvetica-Bold', 7.8)
                c.drawString(x+7*mm, y-2*mm, ln); y -= 10
            if sub:
                for ln in _wl(c, sub, 'Helvetica-Oblique', 7, cw2-8*mm):
                    c.setFillColor(MIDGREY); c.setFont('Helvetica-Oblique', 7)
                    c.drawString(x+7*mm, y-1*mm, ln); y -= 9
            y -= GAP
        y -= 4*mm; return y

    cs = ty - ib_h - 7*mm
    y1 = cs
    for g in grps[:2]: y1 = draw_grp(c, M, y1, g)
    y2 = cs
    for g in grps[2:]: y2 = draw_grp(c, M+cw2+5*mm, y2, g)

    content_bot = min(y1, y2)
    ah3 = 22*mm; ay = content_bot - 8*mm - ah3
    c.setFillColor(TEAL); c.roundRect(M, ay, CW, ah3, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 8.5)
    c.drawString(M+4*mm, ay+16*mm, _u["if_any_no"])
    _tb(c, d["T4_IF_ANY_NO"], M+4*mm, ay+10*mm, CW-8*mm, 'Helvetica', 8, col=WHITE, ld=10)
    _ftr(c, _u["t4_footer"], _u=_u)


# ── Public entry point ────────────────────────────────────────────────────────

_TOOL_DRAW = {'T1': draw_t1, 'T2': draw_t2, 'T3': draw_t3, 'T4': draw_t4}

# Human-readable labels (for UI)
TOOL_LABELS = {
    'T1': 'Tool 1 — Compliance Readiness Self-Check (B1)',
    'T2': 'Tool 2 — Legal Support Decision Tree (B3)',
    'T3': 'Tool 3 — Go public, or stay quiet? (B5)',
    'T4': 'Tool 4 — Do-No-Harm Communications Checklist (B5)',
}


def build_tools_pdf(selected_ids, data=None, buf=None, language="EN"):
    """Generate selected tool pages into a BytesIO buffer.

    Args:
        selected_ids: list/set of tool IDs, e.g. ['T1', 'T3']
        data: dict of {field_key: value} loaded from TOOLS sheet (may include all tools' keys)
        buf: optional existing BytesIO; if None a new one is created.
        language: 'EN', 'FR', or 'ES' — controls structural label translation.
    Returns:
        BytesIO with the tool pages PDF, or None if no tools selected.
    """
    ids = [tid for tid in ['T1', 'T2', 'T3', 'T4'] if tid in (selected_ids or [])]
    if not ids:
        return None
    if buf is None:
        buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    for tid in ids:
        _TOOL_DRAW[tid](c, data, language=language)
        c.showPage()
    c.save()
    buf.seek(0)
    return buf
