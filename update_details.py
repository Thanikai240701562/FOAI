import os

# Update build_full_report.py
with open(r'c:\AI & Web project\build_full_report.py', 'r', encoding='utf-8') as f:
    code = f.read()

code = code.replace("add_title_p(\"2116230701xxx – [CANDIDATE NAME 1]\", size=14, bold=True, space_before=0, space_after=4)\nadd_title_p(\"2116230701xxx – [CANDIDATE NAME 2]\", size=14, bold=True, space_before=0, space_after=4)\nadd_title_p(\"2116230701xxx – [CANDIDATE NAME 3]\", size=14, bold=True, space_before=0, space_after=4)\nadd_title_p(\"2116230701xxx – [CANDIDATE NAME 4]\", size=14, bold=True, space_before=0, space_after=40)",
"""add_title_p("2116240701556 – TARUN ANTONY M", size=14, bold=True, space_before=0, space_after=6)
add_title_p("2116240701562 – THANIKAIVEL S", size=14, bold=True, space_before=0, space_after=40)""")

code = code.replace("[NAME OF THE INSTITUTION / COLLEGE]", "RAJALAKSHMI ENGINEERING COLLEGE")
code = code.replace("[Name of Institution]", "Rajalakshmi Engineering College")
code = code.replace("October, 2026", "NOVEMBER 2026")
code = code.replace('"[NAME OF THE CANDIDATE(S)]" (Register No(s): [2116230701xxx, ...])', 'Tarun Antony M (2116240701556) and Thanikaivel S (2116240701562)')
code = code.replace('[Assistant Professor / Associate Professor]', 'Mrs. B. Deepa, Assistant Professor (SG)')

with open(r'c:\AI & Web project\build_full_report.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Updated build_full_report.py successfully.")
