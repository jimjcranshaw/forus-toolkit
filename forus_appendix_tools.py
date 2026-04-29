"""
forus_appendix_tools.py - Forus Toolkit appendix tool pages (A1-A3)
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


# -- Default text (sourced from TOOLS sheet at runtime) ---

A3_DEFAULTS = {
    "A3_HOW_TO_USE": (
        "Identify which of the four crisis types best describes your situation. Follow that path to "
        "the recommended mechanisms. Contact Forus first - many mechanisms prioritise "
        "network-referred applications over direct applicants."),
    "A3_ROUTING_QUESTION": "What is the primary nature of your funding or resourcing crisis?",
    # Box 2 - Sudden contract loss (Lifeline removed; Lighthouse slot 1, CIVICUS slot 2)
    "A3_BOX2_POINT_1_KEY":   "Lighthouse Global Protection Fund",
    "A3_BOX2_POINT_1_DESC":  "Rapid response for CSOs facing sudden existential crises",
    "A3_BOX2_POINT_2_KEY":   "CIVICUS Response Fund",
    "A3_BOX2_POINT_2_DESC":  "For civic actors globally; member applications prioritised",
    "A3_BOX2_POINT_3_KEY":   "",
    "A3_BOX2_POINT_3_DESC":  "",
    # Box 3 - Regulatory block (Lifeline replaced with Lighthouse)
    "A3_BOX3_POINT_2_KEY":   "Lighthouse Global Protection Fund",
    "A3_BOX3_POINT_2_DESC":  "Covers legal defence costs if under-threat criteria are met",
}
