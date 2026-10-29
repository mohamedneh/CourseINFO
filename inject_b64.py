# -*- coding: utf-8 -*-
import base64
import os

pptx_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\Informatique_Cours_Complet.pptx"
with open(pptx_path, "rb") as f:
    b64_pptx = base64.b64encode(f.read()).decode("utf-8")

# Read template and replace placeholder
template_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\template.html"
with open(template_path, "r", encoding="utf-8") as f:
    content = f.read()

content = content.replace("%%PPTX_BASE64%%", b64_pptx)

index_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\index.html"
with open(index_path, "w", encoding="utf-8") as f:
    f.write(content)
print("Updated index.html")

brain_cinematic = r"C:\Users\Moham\.gemini\antigravity\brain\8e28d94e-0064-4605-ad65-07261bf39388\cinematic_presentation.html"
with open(brain_cinematic, "w", encoding="utf-8") as f:
    f.write(content)
print("Updated cinematic_presentation.html")
