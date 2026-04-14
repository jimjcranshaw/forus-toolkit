"""
forus_appendix_tools.py — Forus Toolkit appendix tool pages (A1–A3)
All text content is read from the TOOLS sheet at runtime via the `data` dict.
draw_a1/a2/a3 each accept a data dict; defaults are the original hardcoded strings.
build_appendix_pdf(selected_ids, data, buf) generates selected pages into a BytesIO.
"""

import io
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas as rl_canvas

TEAL   = HexColor('#00424D'); LBLUE = HexColor('#58C5C7'); PINK  = HexColor('#ED1651')
MINT   = HexColor('#5C9C8E'); LIME  = HexColor('#B2C100'); WHITE = HexColor('#FFFFFF')
OFFWHITE = HexColor('#F4F8F8'); MIDGREY = HexColor('#6B7C80')
LTGREY   = HexColor('#E0EAEA'); DKGREY  = HexColor('#2E4A50')
W, H = A4; M = 20*mm; CW = W - 2*M; CX = W / 2


# ── Default text (sourced from TOOLS sheet at runtime) ───────────────────────

A1_DEFAULTS = {
    "A1_WHY_THIS_MATTERS": (
        "Most crisis failures happen when members expect more than their platform can deliver. "
        "Platforms that have not defined their role invite false expectations — and experience compounded "
        "harm when they cannot deliver on them."),
    "A1_TIER1_TITLE": "TIER 1 — CRISIS COORDINATION NODE",
    "A1_TIER1_DESC": (
        "Your platform directly coordinates emergency response: sourcing legal support, accessing "
        "funds, arranging security or relocation. You are a crisis actor."),
    "A1_TIER1_IND_1": "Dedicated staff with crisis duties protected from other workloads",
    "A1_TIER1_IND_2": "Emergency budget accessible without full board sign-off",
    "A1_TIER1_IND_3": "Secure channel monitored continuously, including nights and weekends",
    "A1_TIER1_IND_4": "Pre-agreed relationships with legal, financial, and security providers",
    "A1_TIER1_IND_5": "Members formally expect direct crisis coordination from you",
    "A1_TIER2_TITLE": "TIER 2 — ACTIVE SOLIDARITY",
    "A1_TIER2_DESC": (
        "Your platform advocates publicly, amplifies member situations, refers to providers, "
        "and presses cases with donors and authorities. A credible, high-value role most "
        "platforms can sustain."),
    "A1_TIER2_IND_1": "Can publish statements on behalf of members within 24 hours",
    "A1_TIER2_IND_2": "Established relationships with embassies, donors, and media contacts",
    "A1_TIER2_IND_3": "Capacity to refer members to specialist support providers",
    "A1_TIER2_IND_4": "Leadership authority to speak publicly on behalf of members in crisis",
    "A1_TIER2_IND_5": "Members understand you advocate and refer — not directly respond",
    "A1_TIER3_TITLE": "TIER 3 — MONITORING & FLAGGING",
    "A1_TIER3_DESC": (
        "Your platform tracks member situations and escalates to Forus or a regional coalition. "
        "You do not coordinate response or advocate publicly — an honest and protective position "
        "in high-exposure contexts."),
    "A1_TIER3_IND_1": "Systems to document member situations and share updates securely",
    "A1_TIER3_IND_2": "Clear escalation pathway to Forus or your regional coalition",
    "A1_TIER3_IND_3": "Members understand your role is to flag and report, not respond",
    "A1_TIER3_IND_4": "Public advocacy would place your platform or leadership at unacceptable risk",
    "A1_OUTCOME_BOX": (
        "A clear statement of your solidarity role prevents false expectations in a crisis. "
        "The B2 section provides model language for each tier. "
        "Revisit annually or whenever your political environment changes significantly."),
}

A2_DEFAULTS = {
    "A2_WHY_THIS_MATTERS": (
        "Diversification pursued from desperation rarely succeeds — and often strains the partnerships "
        "it depends on. This tool checks whether the preconditions for sustainable diversification "
        "are in place before you commit."),
    "A2_GATE_ITEM_1": "Core operations are funded for at least 6 months ahead",
    "A2_GATE_ITEM_2": "Leadership has bandwidth beyond managing current financial pressures",
    "A2_GATE_ITEM_3": "The political environment allows engagement with new donors or partners",
    "A2_GATE_ITEM_4": "This decision is strategy-driven, not desperation-driven",
    "A2_OUTCOME_PROCEED_HEAD": "3–4  ✓   PROCEED TO STEP 2",
    "A2_OUTCOME_PROCEED_SUB":  "Preconditions are in place.",
    "A2_OUTCOME_NOT_YET_HEAD": "FEWER THAN 3  ✓   NOT YET",
    "A2_OUTCOME_NOT_YET_SUB":  "Stabilise first. Revisit in 90 days.",
    "A2_TRUST_ITEM_1":    "Two or more potential partners have a track record of financial collaboration",
    "A2_TRUST_ITEM_2":    "No active disputes or unresolved grievances between potential partners",
    "A2_GOVERNANCE_ITEM_1": "Agreed decision-making authority for the shared arrangement, covering disputes and member exit",
    "A2_GOVERNANCE_ITEM_2": "At least one party has legal capacity to hold shared funds or contracts on behalf of the group",
    "A2_CAPACITY_ITEM_1": "Combined member capacity covers the service or function being mutualised",
    "A2_CAPACITY_ITEM_2": "A lead organisation is willing to absorb the initial administrative burden",
    "A2_OUTCOME_ALL_6_HEAD": "ALL 6 TICKED  →  PROCEED TO MODEL SELECTION — see Part 6 mechanism directory in the Forus Toolkit app",
    "A2_OUTCOME_ALL_6_BODY": (
        "Fewer than 6: address gaps before committing. Governance gaps are the most common cause of "
        "mutualisation failure — resolve these first, then trust, then capacity."),
}

A3_DEFAULTS = {
    "A3_HOW_TO_USE": (
        "Identify which of the four crisis types best describes your situation. Follow that path to "
        "the recommended mechanisms. Contact Forus first — many mechanisms prioritise "
        "network-referred applications over direct applicants."),
    "A3_ROUTING_QUESTION": "What is the primary nature of your funding or resourcing crisis?",
    # Box 1 — Individual / Leader at risk
    "A3_BOX1_TITLE":         "INDIVIDUAL OR LEADER AT IMMEDIATE RISK",
    "A3_BOX1_SUB":           "A staff member is detained, threatened, or forced to flee.",
    "A3_BOX1_POINT_1_KEY":   "Frontline Defenders",
    "A3_BOX1_POINT_1_DESC":  "Emergency support for human rights defenders at risk — frontlinedefenders.org",
    "A3_BOX1_POINT_2_KEY":   "OMCT Emergency Fund",
    "A3_BOX1_POINT_2_DESC":  "For organisations defending HRDs under persecution",
    "A3_BOX1_NOTE":          "This is primarily a security crisis. Go to Legal Support Decision Tree (Tool 2) first. Emergency funding follows legal navigation, not the other way around.",
    # Box 2 — Sudden contract loss
    "A3_BOX2_TITLE":         "SUDDEN CONTRACT LOSS OR FUNDING GAP",
    "A3_BOX2_SUB":           "A major funder exits representing 30% or more of your budget.",
    "A3_BOX2_POINT_1_KEY":   "Lifeline (Embattled CSO Fund)",
    "A3_BOX2_POINT_1_DESC":  "Multi-donor emergency grants for CSOs under threat — apply early",
    "A3_BOX2_POINT_2_KEY":   "Lighthouse Global Protection Fund",
    "A3_BOX2_POINT_2_DESC":  "Rapid response for CSOs facing sudden existential crises",
    "A3_BOX2_POINT_3_KEY":   "CIVICUS Response Fund",
    "A3_BOX2_POINT_3_DESC":  "For civic actors globally; member applications prioritised",
    "A3_BOX2_NOTE":          "Apply in parallel, not in sequence. Most have 2–4 week application-to-disbursement timelines.",
    # Box 3 — Regulatory block
    "A3_BOX3_TITLE":         "REGULATORY BLOCK OR ACCOUNT FREEZE",
    "A3_BOX3_SUB":           "Operations legally restricted, accounts blocked, or registration under challenge.",
    "A3_BOX3_POINT_1_KEY":   "Legal navigation first",
    "A3_BOX3_POINT_1_DESC":  "See Legal Support Decision Tree (Tool 2) — legal crisis precedes funding crisis",
    "A3_BOX3_POINT_2_KEY":   "Lifeline / Lighthouse",
    "A3_BOX3_POINT_2_DESC":  "Both cover legal defence costs if under-threat criteria are met",
    "A3_BOX3_POINT_3_KEY":   "Escalate to Forus immediately",
    "A3_BOX3_POINT_3_DESC":  "Regulatory crises have hard legal deadlines. Network solidarity advocacy can unlock informal bridge support.",
    # Box 4 — Sector-wide contraction
    "A3_BOX4_TITLE":         "SECTOR-WIDE FUNDING CONTRACTION",
    "A3_BOX4_SUB":           "Broad funding cuts affecting multiple organisations simultaneously.",
    "A3_BOX4_POINT_1_KEY":   "Building Responses Together (BRT)",
    "A3_BOX4_POINT_1_DESC":  "Referral and coordination support hosted by Global Focus (DANIDA-funded — verify current status)",
    "A3_BOX4_POINT_2_KEY":   "Diversification pathway",
    "A3_BOX4_POINT_2_DESC":  "See Diversification Readiness Gate (Appendix A2) before committing to new income models",
    "A3_BOX4_NOTE":          "This is a structural challenge. Emergency mechanisms are not designed for sector-wide contraction — strategy is the right response, not emergency applications.",
}


# ── Localised UI strings (structural labels, headings) ───────────────────────

_UI = {
    "EN": {
        "appendix_tag":    "APPENDIX {num}  ·  {sec}",
        "toolkit_name":    "FORUS RESILIENCE TOOLKIT",
        "footer_url":      "forus-international.org",
        "why_matters":     "WHY THIS MATTERS",
        "how_to_use_a1":   "HOW TO USE:  Read each tier. Identify which best matches your platform today. Score one point per tick. Communicate your tier to members in writing.",
        "signs_your_role": "SIGNS THIS IS YOUR ROLE",
        "once_identified": "ONCE YOU HAVE IDENTIFIED YOUR TIER — COMMUNICATE IT TO MEMBERS IN WRITING",
        "a1_footer":       "Appendix A1 of 3  ·  Linked from B2 — Solidarity Mechanisms",
        "step1_lbl":       "STEP 1  —  THE READINESS GATE",
        "step1_prompt":    "Tick each statement that is true for your organisation today:",
        "step2_lbl":       "STEP 2  —  MUTUALISATION READINESS  (complete only if proceeding)",
        "step2_prompt":    "Tick each item. A single gap should trigger a planning conversation before you commit.",
        "trust":           "TRUST",
        "governance":      "GOVERNANCE",
        "capacity":        "CAPACITY",
        "a2_footer":       "Appendix A2 of 3  ·  Part 6 — Diversification & Mutualisation",
        "how_to_use_a3":   "HOW TO USE",
        "note_lbl":        "NOTE",
        "a3_footer":       "Appendix A3 of 3  ·  Part 4 — Emergency Funding",
    },
    "FR": {
        "appendix_tag":    "ANNEXE {num}  ·  {sec}",
        "toolkit_name":    "BOITE A OUTILS RESILIENCE FORUS",
        "footer_url":      "forus-international.org",
        "why_matters":     "POURQUOI C'EST IMPORTANT",
        "how_to_use_a1":   "MODE D'EMPLOI :  Lisez chaque niveau. Identifiez celui qui correspond le mieux a votre plateforme. Un point par coche. Communiquez votre niveau aux membres par ecrit.",
        "signs_your_role": "SIGNES QUE C'EST VOTRE ROLE",
        "once_identified": "UNE FOIS VOTRE NIVEAU IDENTIFIE — COMMUNIQUEZ-LE AUX MEMBRES PAR ECRIT",
        "a1_footer":       "Annexe A1 sur 3  ·  Lie a B2 — Mecanismes de solidarite",
        "step1_lbl":       "ETAPE 1  —  EVALUATION DE LA DISPONIBILITE",
        "step1_prompt":    "Cochez chaque affirmation vraie pour votre organisation aujourd'hui :",
        "step2_lbl":       "ETAPE 2  —  MUTUALISATION  (a completer uniquement si vous procedez)",
        "step2_prompt":    "Cochez chaque element. Une lacune devrait declencher une discussion avant de s'engager.",
        "trust":           "CONFIANCE",
        "governance":      "GOUVERNANCE",
        "capacity":        "CAPACITE",
        "a2_footer":       "Annexe A2 sur 3  ·  Partie 6 — Diversification et mutualisation",
        "how_to_use_a3":   "MODE D'EMPLOI",
        "note_lbl":        "NOTE",
        "a3_footer":       "Annexe A3 sur 3  ·  Partie 4 — Financement d'urgence",
    },
    "ES": {
        "appendix_tag":    "ANEXO {num}  ·  {sec}",
        "toolkit_name":    "KIT DE HERRAMIENTAS FORUS",
        "footer_url":      "forus-international.org",
        "why_matters":     "POR QUE ES IMPORTANTE",
        "how_to_use_a1":   "COMO USAR:  Lea cada nivel. Identifique cual se adapta mejor a su plataforma. Un punto por marca. Comunique su nivel a los miembros por escrito.",
        "signs_your_role": "INDICIOS DE QUE ESTE ES SU ROL",
        "once_identified": "UNA VEZ IDENTIFICADO SU NIVEL — COMUNIQUESELO A LOS MIEMBROS POR ESCRITO",
        "a1_footer":       "Anexo A1 de 3  ·  Vinculado a B2 — Mecanismos de solidaridad",
        "step1_lbl":       "PASO 1  —  EVALUACION DE DISPONIBILIDAD",
        "step1_prompt":    "Marque cada afirmacion verdadera para su organizacion hoy:",
        "step2_lbl":       "PASO 2  —  DISPONIBILIDAD PARA LA MUTUALIZACION  (solo si procede)",
        "step2_prompt":    "Marque cada elemento. Una brecha debe generar una conversacion antes de comprometerse.",
        "trust":           "CONFIANZA",
        "governance":      "GOBERNANZA",
        "capacity":        "CAPACIDAD",
        "a2_footer":       "Anexo A2 de 3  ·  Parte 6 — Diversificacion y mutualizacion",
        "how_to_use_a3":   "COMO USAR",
        "note_lbl":        "NOTA",
        "a3_footer":       "Anexo A3 de 3  ·  Parte 4 — Financiacion de emergencia",
    },
}


# ── Drawing utilities ─────────────────────────────────────────────────────────

def _hdr(c, num, sec, title, sub=None, _u=None):
    if _u is None: _u = _UI["EN"]
    bh = 48*mm; c.setFillColor(TEAL); c.rect(0, H-bh, W, bh, fill=1, stroke=0)
    tag = _u["appendix_tag"].format(num=num, sec=sec)
    tw = c.stringWidth(tag, 'Helvetica-Bold', 7) + 6*mm
    c.setFillColor(LBLUE); c.roundRect(M, H-11*mm, tw, 7*mm, 1*mm, fill=1, stroke=0)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 7); c.drawString(M+3*mm, H-7.5*mm, tag)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 20); c.drawString(M, H-27*mm, title)
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


def _gh(c, x, y, w, title, col=TEAL):
    c.setFillColor(col); c.roundRect(x, y-5.5*mm, w, 5.5*mm, 1*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 7); c.drawString(x+3*mm, y-3.8*mm, title)


def _cbx(c, x, y, sz=3.8*mm, col=TEAL):
    c.setStrokeColor(col); c.setLineWidth(0.8); c.rect(x, y, sz, sz, fill=0, stroke=1)


def _draw_sect(c, x, y, title, col, items, cw_):
    _gh(c, x, y, cw_, title, col=col); y -= 8*mm
    for item in items:
        sz = 3.8*mm; _cbx(c, x+1*mm, y-sz-0.5*mm, sz, col=col)
        lns = _wl(c, item, 'Helvetica', 8, cw_-8*mm)
        for ln in lns:
            c.setFillColor(DKGREY); c.setFont('Helvetica', 8)
            c.drawString(x+7*mm, y-2*mm, ln); y -= 10
        y -= 5
    y -= 4*mm; return y


# ── Appendix A1 ───────────────────────────────────────────────────────────────

def draw_a1(c, data=None, language="EN"):
    d = dict(A1_DEFAULTS)
    if data: d.update(data)
    _u = _UI.get(language.upper(), _UI["EN"])

    _hdr(c, 'A1 of 3', 'B2 · SOLIDARITY MECHANISMS',
         "What is your platform's solidarity role?",
         'Complete with your leadership team before a crisis — not during one', _u=_u)
    ty = H - 48*mm - 6*mm

    ib_h = 20*mm
    c.setFillColor(OFFWHITE); c.roundRect(M, ty-ib_h, CW, ib_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 8.5)
    c.drawString(M+4*mm, ty-6*mm, _u["why_matters"])
    _tb(c, d["A1_WHY_THIS_MATTERS"], M+4*mm, ty-13*mm, CW-8*mm, 'Helvetica', 8, col=MIDGREY, ld=10.5)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 7.5)
    c.drawString(M, ty-ib_h-5*mm, _u["how_to_use_a1"])

    gap = 5*mm; ncols = 3; cw = (CW - (ncols-1)*gap) / ncols
    col_xs = [M + i*(cw+gap) for i in range(ncols)]
    cs = ty - ib_h - 14*mm

    tiers = [
        (TEAL,   d["A1_TIER1_TITLE"], d["A1_TIER1_DESC"],
         [d[f"A1_TIER1_IND_{i}"] for i in range(1, 6)]),
        (MINT,   d["A1_TIER2_TITLE"], d["A1_TIER2_DESC"],
         [d[f"A1_TIER2_IND_{i}"] for i in range(1, 6)]),
        (DKGREY, d["A1_TIER3_TITLE"], d["A1_TIER3_DESC"],
         [d[f"A1_TIER3_IND_{i}"] for i in range(1, 5)]),
    ]

    col_bottoms = []
    for col_x, (tcol, ttitle, tdesc, tinds) in zip(col_xs, tiers):
        y = cs
        _gh(c, col_x, y, cw, ttitle, col=tcol); y -= 8*mm
        y = _tb(c, tdesc, col_x+2*mm, y, cw-4*mm, 'Helvetica', 7.5, col=MIDGREY, ld=10)
        y -= 5*mm
        c.setFillColor(tcol); c.setFont('Helvetica-Bold', 6.5)
        c.drawString(col_x+2*mm, y, _u["signs_your_role"]); y -= 5*mm
        for ind in tinds:
            sz = 3.5*mm; _cbx(c, col_x+1*mm, y-sz-0.5*mm, sz, col=tcol)
            lns = _wl(c, ind, 'Helvetica', 7.5, cw-7*mm)
            for ln in lns:
                c.setFillColor(DKGREY); c.setFont('Helvetica', 7.5)
                c.drawString(col_x+6*mm, y-1.5*mm, ln); y -= 10
            y -= 5
        col_bottoms.append(y)

    content_bot = min(col_bottoms)
    ab_h = 22*mm; ab_y = content_bot - 8*mm - ab_h
    c.setFillColor(TEAL); c.roundRect(M, ab_y, CW, ab_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 8.5)
    c.drawString(M+4*mm, ab_y+16*mm, _u["once_identified"])
    _tb(c, d["A1_OUTCOME_BOX"], M+4*mm, ab_y+10*mm, CW-8*mm, 'Helvetica', 8, col=WHITE, ld=10)
    _ftr(c, _u["a1_footer"], _u=_u)


# ── Appendix A2 ───────────────────────────────────────────────────────────────

def draw_a2(c, data=None, language="EN"):
    d = dict(A2_DEFAULTS)
    if data: d.update(data)
    _u = _UI.get(language.upper(), _UI["EN"])

    _hdr(c, 'A2 of 3', 'PART 6 · DIVERSIFICATION & MUTUALISATION',
         'Diversification Readiness Gate',
         'Complete before committing to a new income model or mutualisation arrangement', _u=_u)
    ty = H - 48*mm - 6*mm

    ib_h = 20*mm
    c.setFillColor(OFFWHITE); c.roundRect(M, ty-ib_h, CW, ib_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 8.5)
    c.drawString(M+4*mm, ty-6*mm, _u["why_matters"])
    _tb(c, d["A2_WHY_THIS_MATTERS"], M+4*mm, ty-13*mm, CW-8*mm, 'Helvetica', 8, col=MIDGREY, ld=10.5)

    # Step 1
    s1y = ty - ib_h - 9*mm
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 9)
    c.drawString(M, s1y, _u["step1_lbl"])

    s1bh = 46*mm; s1by = s1y - 6*mm - s1bh
    c.setFillColor(OFFWHITE); c.roundRect(M, s1by, CW, s1bh, 2*mm, fill=1, stroke=0)
    c.setFillColor(DKGREY); c.setFont('Helvetica', 8)
    c.drawString(M+4*mm, s1by+s1bh-7*mm, _u["step1_prompt"])

    gate_items = [d[f"A2_GATE_ITEM_{i}"] for i in range(1, 5)]
    gy = s1by + s1bh - 14*mm
    for item in gate_items:
        sz = 3.8*mm; _cbx(c, M+4*mm, gy-sz-0.5*mm, sz, col=TEAL)
        c.setFillColor(DKGREY); c.setFont('Helvetica', 8)
        c.drawString(M+11*mm, gy-2*mm, item); gy -= 17

    ob_w = (CW - 5*mm) / 2; ob_h = 14*mm; ob_y = s1by - 5*mm - ob_h
    c.setFillColor(MINT); c.roundRect(M, ob_y, ob_w, ob_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 8.5)
    c.drawCentredString(M+ob_w/2, ob_y+8.5*mm, d["A2_OUTCOME_PROCEED_HEAD"])
    c.setFont('Helvetica', 7.5)
    c.drawCentredString(M+ob_w/2, ob_y+3.5*mm, d["A2_OUTCOME_PROCEED_SUB"])

    c.setFillColor(PINK); c.roundRect(M+ob_w+5*mm, ob_y, ob_w, ob_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 8.5)
    c.drawCentredString(M+ob_w+5*mm+ob_w/2, ob_y+8.5*mm, d["A2_OUTCOME_NOT_YET_HEAD"])
    c.setFont('Helvetica', 7.5)
    c.drawCentredString(M+ob_w+5*mm+ob_w/2, ob_y+3.5*mm, d["A2_OUTCOME_NOT_YET_SUB"])

    # Step 2
    s2y = ob_y - 9*mm
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 9)
    c.drawString(M, s2y, _u["step2_lbl"])
    c.setFillColor(MIDGREY); c.setFont('Helvetica', 7.5)
    c.drawString(M, s2y-6*mm, _u["step2_prompt"])

    cw2 = (CW - 5*mm) / 2; s2cs = s2y - 14*mm
    ly = _draw_sect(c, M, s2cs, _u["trust"], TEAL,
        [d["A2_TRUST_ITEM_1"], d["A2_TRUST_ITEM_2"]], cw2)
    ly = _draw_sect(c, M, ly, _u["governance"], MINT,
        [d["A2_GOVERNANCE_ITEM_1"], d["A2_GOVERNANCE_ITEM_2"]], cw2)
    ry = _draw_sect(c, M+cw2+5*mm, s2cs, _u["capacity"], DKGREY,
        [d["A2_CAPACITY_ITEM_1"], d["A2_CAPACITY_ITEM_2"]], cw2)

    content_bot = min(ly, ry)
    ab_h = 22*mm; ab_y = content_bot - 8*mm - ab_h
    c.setFillColor(TEAL); c.roundRect(M, ab_y, CW, ab_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 8.5)
    c.drawString(M+4*mm, ab_y+16*mm, d["A2_OUTCOME_ALL_6_HEAD"])
    _tb(c, d["A2_OUTCOME_ALL_6_BODY"], M+4*mm, ab_y+10*mm, CW-8*mm, 'Helvetica', 8, col=WHITE, ld=10)
    _ftr(c, _u["a2_footer"], _u=_u)


# ── Appendix A3 ───────────────────────────────────────────────────────────────

def draw_a3(c, data=None, language="EN"):
    d = dict(A3_DEFAULTS)
    if data: d.update(data)
    _u = _UI.get(language.upper(), _UI["EN"])

    _hdr(c, 'A3 of 3', 'PART 4 · EMERGENCY FUNDING',
         'Emergency Funding Navigator',
         'Match your crisis type to the most accessible rapid-response mechanisms', _u=_u)
    ty = H - 48*mm - 6*mm

    ib_h = 18*mm
    c.setFillColor(OFFWHITE); c.roundRect(M, ty-ib_h, CW, ib_h, 2*mm, fill=1, stroke=0)
    c.setFillColor(TEAL); c.setFont('Helvetica-Bold', 8.5)
    c.drawString(M+4*mm, ty-6*mm, _u["how_to_use_a3"])
    _tb(c, d["A3_HOW_TO_USE"], M+4*mm, ty-12*mm, CW-8*mm, 'Helvetica', 8, col=MIDGREY, ld=10)

    rqh = 10*mm; rqy = ty - ib_h - 6*mm
    c.setFillColor(DKGREY); c.roundRect(M, rqy-rqh, CW, rqh, 2*mm, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 9)
    c.drawCentredString(CX, rqy-4.5*mm, d["A3_ROUTING_QUESTION"])

    bw = (CW - 5*mm) / 2; bh = 65*mm
    grid_top = rqy - rqh - 5*mm

    boxes = [
        (M,         grid_top,          TEAL,   1),
        (M+bw+5*mm, grid_top,          PINK,   2),
        (M,         grid_top-bh-5*mm,  DKGREY, 3),
        (M+bw+5*mm, grid_top-bh-5*mm,  MINT,   4),
    ]

    for bx, by_top, bcol, n in boxes:
        box_y = by_top - bh
        c.setFillColor(bcol); c.roundRect(bx, box_y, bw, bh, 2*mm, fill=1, stroke=0)
        c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 7.5)
        c.drawString(bx+3*mm, by_top-7*mm, d[f"A3_BOX{n}_TITLE"])
        y = by_top - 13*mm
        y = _tb(c, d[f"A3_BOX{n}_SUB"], bx+3*mm, y, bw-6*mm, 'Helvetica-Oblique', 7.5, col=WHITE, ld=10)
        y -= 3*mm

        # Collect points for this box
        pts = []
        for p in range(1, 5):
            key_k = f"A3_BOX{n}_POINT_{p}_KEY"
            key_v = f"A3_BOX{n}_POINT_{p}_DESC"
            if key_k in d:
                pts.append((d[key_k], d[key_v]))
        note_key = f"A3_BOX{n}_NOTE"

        for bkey, bval in pts:
            c.setFillColor(WHITE); c.setFont('Helvetica-Bold', 7.5)
            c.drawString(bx+3*mm, y, '·  ' + bkey); y -= 10
            y = _tb(c, bval, bx+5*mm, y, bw-8*mm, 'Helvetica', 7, col=WHITE, ld=9); y -= 3*mm

        if note_key in d and d[note_key]:
            c.setFillColor(WHITE); c.setFont('Helvetica-BoldOblique', 6.5)
            c.drawString(bx+3*mm, y, _u["note_lbl"]); y -= 9
            _tb(c, d[note_key], bx+3*mm, y, bw-6*mm, 'Helvetica-Oblique', 6.5, col=WHITE, ld=9)

    _ftr(c, _u["a3_footer"], _u=_u)


# ── Public entry point ────────────────────────────────────────────────────────

_APPENDIX_DRAW = {'A1': draw_a1, 'A2': draw_a2, 'A3': draw_a3}

APPENDIX_LABELS = {
    'A1': 'Appendix A1 — Platform Role Clarifier (B2)',
    'A2': 'Appendix A2 — Diversification Readiness Gate (B6)',
    'A3': 'Appendix A3 — Emergency Funding Navigator (B4)',
}


def build_appendix_pdf(selected_ids, data=None, buf=None, language="EN"):
    """Generate selected appendix tool pages into a BytesIO buffer.

    Args:
        selected_ids: list/set of appendix IDs, e.g. ['A1', 'A3']
        data: dict of {field_key: value} loaded from TOOLS sheet
        buf: optional existing BytesIO; if None a new one is created.
        language: 'EN', 'FR', or 'ES' — controls structural label translation.
    Returns:
        BytesIO with the appendix pages PDF, or None if no appendices selected.
    """
    ids = [tid for tid in ['A1', 'A2', 'A3'] if tid in (selected_ids or [])]
    if not ids:
        return None
    if buf is None:
        buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    for tid in ids:
        _APPENDIX_DRAW[tid](c, data, language=language)
        c.showPage()
    c.save()
    buf.seek(0)
    return buf
