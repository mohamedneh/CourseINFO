# -*- coding: utf-8 -*-
"""
Script to build the comprehensive 26-slide index.html with:
1. Auto-collapsing / hide-on-click top header bar.
2. The interactive Windows 11 Desktop Simulator slide (Slide 13).
3. The 5 System Icons deep academic explanation slide (Slide 14).
4. The Taskbar & Power States (Sleep vs Hibernate) slide (Slide 15).
5. All 26 slides fully articulated with read-aloud boxes for Mr. Nehmar Mohamed.
6. Embedded Base64 for the 26-slide PPTX file.
"""

import base64
import os

pptx_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\Informatique_Cours_Complet.pptx"
with open(pptx_path, "rb") as f:
    b64_pptx = base64.b64encode(f.read()).decode("utf-8")

print(f"Loaded PPTX base64, length: {len(b64_pptx)}")
