import os

path = r'c:\Users\ARuiz\OneDrive - CIEE\scripts\GeneradorContratos\generador_contratos.py'
with open(path, 'r', encoding='utf-8') as f:
    text = f.read()

# Replace all font usages
text = text.replace('"System"', '"Segoe UI"')
# Update header color
text = text.replace('fg_color=("gray85", "gray20")', 'fg_color=("#1f538d", "#14375e")')
text = text.replace('text_color="gray"', 'text_color=("gray60", "gray50")')

with open(path, 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated generador_contratos.py")
