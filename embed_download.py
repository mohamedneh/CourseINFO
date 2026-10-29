# -*- coding: utf-8 -*-
import base64
import os

pptx_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\Informatique_Cours_Complet.pptx"
with open(pptx_path, "rb") as f:
    b64_pptx = base64.b64encode(f.read()).decode("utf-8")

print(f"Base64 length: {len(b64_pptx)}")

# Update index.html
index_path = r"C:\Users\Moham\.gemini\antigravity\scratch\informatique_presentation\index.html"
with open(index_path, "r", encoding="utf-8") as f:
    html_content = f.read()

# Replace the download link with a JS call
old_link = '<a href="Informatique_Cours_Complet.pptx" download'
new_button = '<button onclick="downloadPPTXFile()"'

if old_link in html_content:
    html_content = html_content.replace(old_link, new_button)
    # Also replace closing tag </a> with </button> for this specific button
    html_content = html_content.replace('<span>📊 تحميل ملف PowerPoint (.pptx)</span>\n              </a>', '<span>📊 تحميل ملف PowerPoint (.pptx)</span>\n              </button>')

# Add JS download function with embedded base64
download_js = f"""
    // Embedded PPTX Downloader
    const PPTX_BASE64 = "{b64_pptx}";
    function downloadPPTXFile() {{
      try {{
        const binaryString = window.atob(PPTX_BASE64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {{
          bytes[i] = binaryString.charCodeAt(i);
        }}
        const blob = new Blob([bytes], {{ type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }});
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'Informatique_Cours_Complet.pptx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        playSound('fanfare');
        alert('✅ بدأ تحميل ملف PowerPoint بنجاح إلى مجلد التنزيلات (Downloads)!');
      }} catch (err) {{
        alert('حدث خطأ أثناء فك التشفير: ' + err.message);
      }}
    }}
"""

if "function downloadPPTXFile()" not in html_content:
    html_content = html_content.replace("</script>\n</body>", download_js + "\n  </script>\n</body>")

with open(index_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print("Updated index.html successfully.")

# Also update cinematic_presentation.html in brain
brain_cinematic = r"C:\Users\Moham\.gemini\antigravity\brain\8e28d94e-0064-4605-ad65-07261bf39388\cinematic_presentation.html"
with open(brain_cinematic, "w", encoding="utf-8") as f:
    f.write(html_content)
print("Updated cinematic_presentation.html successfully.")
