#!/usr/bin/env python3
"""
Patches <a:buChar> elements in each slide with a matching <a:buClr> based on
a JSON queue (one array per slide, each entry {count, color} in call order).
Run AFTER pptxgenjs writes the pptx and BEFORE any further packaging steps.
"""
import sys, json, re, zipfile, shutil, os

def patch(pptx_path, queue_path, out_path):
    with open(queue_path) as f:
        queue = json.load(f)

    tmpdir = pptx_path + '_patchtmp'
    if os.path.exists(tmpdir):
        shutil.rmtree(tmpdir)
    os.makedirs(tmpdir)
    with zipfile.ZipFile(pptx_path) as z:
        z.extractall(tmpdir)

    # Order in schema: pPr -> (lnSpc,spcBef,spcAft)* -> buClr -> buSzPct -> buChar
    # Capture the pPr-open + spacing prefix separately from the buSzPct+buChar suffix,
    # so buClr can be inserted between them in valid schema order.
    buchar_re = re.compile(
        r'(<a:pPr[^>]*>(?:<a:lnSpc>.*?</a:lnSpc>|<a:spcBef>.*?</a:spcBef>|<a:spcAft>.*?</a:spcAft>)*)'
        r'(<a:buSzPct[^/]*/>)?(<a:buChar char="[^"]*"/>)',
        re.DOTALL
    )

    for i, slide_log in enumerate(queue, start=1):
        slide_path = os.path.join(tmpdir, 'ppt', 'slides', f'slide{i}.xml')
        if not os.path.exists(slide_path) or not slide_log:
            continue
        with open(slide_path, encoding='utf-8') as f:
            xml = f.read()

        # Build flat list of colors, one per buChar occurrence, in order
        colors = []
        for entry in slide_log:
            colors.extend([entry['color']] * entry['count'])

        matches = list(buchar_re.finditer(xml))
        if len(matches) != len(colors):
            print(f'WARNING slide{i}: buChar count {len(matches)} != expected {len(colors)}', file=sys.stderr)

        # Rebuild string, inserting buClr before buChar for each match
        out = []
        last_end = 0
        for idx, m in enumerate(matches):
            color = colors[idx] if idx < len(colors) else None
            out.append(xml[last_end:m.start()])
            pPr_open = m.group(1)
            buSzPct = m.group(2) or ''
            buChar = m.group(3)
            out.append(pPr_open)
            if color:
                out.append(f'<a:buClr><a:srgbClr val="{color}"/></a:buClr>')
            out.append(buSzPct)
            out.append(buChar)
            last_end = m.end()
        out.append(xml[last_end:])
        xml = ''.join(out)

        with open(slide_path, 'w', encoding='utf-8') as f:
            f.write(xml)

    # rezip
    if os.path.exists(out_path):
        os.remove(out_path)
    zf = zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk(tmpdir):
        for file in files:
            full = os.path.join(root, file)
            rel = os.path.relpath(full, tmpdir)
            zf.write(full, rel)
    zf.close()
    shutil.rmtree(tmpdir)
    print(f'Patched -> {out_path}')

if __name__ == '__main__':
    patch(sys.argv[1], sys.argv[2], sys.argv[3])
