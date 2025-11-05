import zipfile
import sys
from pathlib import Path

TEMPLATE = Path(__file__).parents[1] / 'templates' / 'New_Template.docx'
PLACEHOLDERS = ['[NO_TEST]', '[REV]', '[DATE]']

if not TEMPLATE.exists():
    print(f"Template not found at: {TEMPLATE}")
    sys.exit(1)

with zipfile.ZipFile(TEMPLATE, 'r') as z:
    found = False
    for name in z.namelist():
        if not name.endswith('.xml'):
            continue
        data = z.read(name).decode('utf-8', errors='ignore')
        for ph in PLACEHOLDERS:
            if ph in data:
                found = True
                idx = data.find(ph)
                start = max(0, idx-60)
                end = min(len(data), idx+60)
                snippet = data[start:end].replace('\n', '')
                print(f"FOUND {ph} in {name}: ...{snippet}...")
    if not found:
        print('No placeholders found in any xml parts.')
