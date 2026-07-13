"""
Post-process pandoc-generated docx to meet WUST 2024 format requirements.
Source: 武汉科技大学本科毕业设计(论文)文件汇编(2024版) 武科大教发〔2024〕39号.

Handles: page setup, headers/footers, text justification, table 三线制,
table/figure captions, heading styles, math font.
"""
import zipfile, os, shutil, io, copy, subprocess
from lxml import etree

# ── Config ──────────────────────────────────────────────────
MD_FILE = '/home/hyhml/my-notes/毕业论文.md'
REFERENCE_DOCX = '/home/hyhml/my-notes/reference.docx'
OUTPUT = '/home/hyhml/my-notes/毕业论文初稿.docx'
HEADER_TEXT = '武汉科技大学本科毕业论文'

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
V_NS = 'urn:schemas-microsoft-com:vml'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

W = f'{{{W_NS}}}'
M = f'{{{M_NS}}}'
R = f'{{{R_NS}}}'
V = f'{{{V_NS}}}'
WP = f'{{{WP_NS}}}'
PIC = f'{{{PIC_NS}}}'

# ── Step 1: Fix reference.docx styles ────────────────────────
def fix_reference_styles():
    """Fix known issues in reference.docx styles.
    - All headings: strip theme fonts, use explicit TNR (western) + 黑体 (CJK)
    - Heading 4: remove theme color + italic, set 黑体 小四号 bold
    - Add 三线制 table style
    """
    print("Step 1: Fixing reference.docx styles...")
    tmp_dir = '/tmp/ref_fix'
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    with zipfile.ZipFile(REFERENCE_DOCX, 'r') as z:
        z.extractall(tmp_dir)

    styles_path = os.path.join(tmp_dir, 'word', 'styles.xml')
    styles_tree = etree.parse(styles_path)
    styles_root = styles_tree.getroot()

    # --- Fix Normal style: set firstLine to 480 DXA (2 chars at 小四号) ---
    for style in styles_root.iter(f'{W}style'):
        style_id = style.get(f'{W}styleId')
        if style_id == 'Normal':
            pPr = style.find(f'{W}pPr')
            if pPr is not None:
                ind = pPr.find(f'{W}ind')
                if ind is not None:
                    ind.set(f'{W}firstLineChars', '200')   # 2 characters
                    # Remove DXA-based firstLine
                    try:
                        del ind.attrib[f'{{{W_NS}}}firstLine']
                    except KeyError:
                        pass
            print("  Fixed Normal style: firstLineChars=200 (2-char indent)")

    # --- Fix ALL heading styles: no bold, no indent, explicit fonts ---
    heading_ids = ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5']
    for style in styles_root.iter(f'{W}style'):
        style_id = style.get(f'{W}styleId')
        if style_id not in heading_ids:
            continue
        # Remove bold (b and bCs) — WUST headings are not bold
        rPr = style.find(f'{W}rPr')
        if rPr is not None:
            for b_tag in ['b', 'bCs']:
                b = rPr.find(f'{W}{b_tag}')
                if b is not None:
                    rPr.remove(b)
            rFonts = rPr.find(f'{W}rFonts')
            if rFonts is not None:
                # Strip theme font references that can override explicit fonts
                for attr in ['asciiTheme', 'eastAsiaTheme', 'hAnsiTheme', 'cstheme']:
                    try:
                        del rFonts.attrib[f'{{{W_NS}}}{attr}']
                    except KeyError:
                        pass
                # Set explicit fonts: TNR for western, 黑体 for CJK
                rFonts.set(f'{W}ascii', 'Times New Roman')
                rFonts.set(f'{W}hAnsi', 'Times New Roman')
                rFonts.set(f'{W}eastAsia', '黑体')
                rFonts.set(f'{W}cs', 'Times New Roman')
        # Explicitly set firstLine=0 to override Normal style inheritance
        pPr = style.find(f'{W}pPr')
        if pPr is None:
            pPr = etree.SubElement(style, f'{W}pPr')
        ind = pPr.find(f'{W}ind')
        if ind is None:
            ind = etree.SubElement(pPr, f'{W}ind')
        ind.set(f'{W}firstLine', '0')
        ind.set(f'{W}firstLineChars', '0')   # override Normal's 200

    # --- Fix BodyText: remove 6pt after-spacing ---
    for style in styles_root.iter(f'{W}style'):
        style_id = style.get(f'{W}styleId')
        if style_id == 'BodyText':
            pPr = style.find(f'{W}pPr')
            if pPr is not None:
                spacing = pPr.find(f'{W}spacing')
                if spacing is not None:
                    spacing.set(f'{W}after', '0')
            print("  Fixed BodyText: after=0 (no 6pt gap)")

    # --- Disable widow/orphan control in all styles ---
    for style in styles_root.iter(f'{W}style'):
        pPr = style.find(f'{W}pPr')
        if pPr is not None:
            wc = pPr.find(f'{W}widowControl')
            if wc is not None:
                wc.set(f'{W}val', 'false')
            else:
                etree.SubElement(pPr, f'{W}widowControl', {f'{W}val': 'false'})

    # --- Fix heading styles: remove after-spacing ---
    for style in styles_root.iter(f'{W}style'):
        style_id = style.get(f'{W}styleId')
        if style_id in heading_ids:
            pPr = style.find(f'{W}pPr')
            if pPr is not None:
                spacing = pPr.find(f'{W}spacing')
                if spacing is not None:
                    spacing.set(f'{W}after', '0')
            # Also remove H1 before spacing — let paragraph-level override handle it
            if style_id == 'Heading1' and pPr is not None:
                spacing = pPr.find(f'{W}spacing')
                if spacing is not None:
                    spacing.set(f'{W}before', '0')   # para-level sets 300

    # --- Heading 4 specific: remove color + italic, fix size ---
    for style in styles_root.iter(f'{W}style'):
        style_id = style.get(f'{W}styleId')
        if style_id == 'Heading4':
            rPr = style.find(f'{W}rPr')
            if rPr is not None:
                color = rPr.find(f'{W}color')
                if color is not None:
                    rPr.remove(color)
                i = rPr.find(f'{W}i')
                if i is not None:
                    rPr.remove(i)
                iCs = rPr.find(f'{W}iCs')
                if iCs is not None:
                    rPr.remove(iCs)
                sz = rPr.find(f'{W}sz')
                if sz is not None:
                    sz.set(f'{W}val', '24')
            pPr = style.find(f'{W}pPr')
            if pPr is not None:
                spacing = pPr.find(f'{W}spacing')
                if spacing is not None:
                    spacing.set(f'{W}before', '120')
                    spacing.set(f'{W}after', '0')
            print("  Fixed Heading 4: removed color/italic, set 黑体 小四号")

    # --- Heading 5 specific: remove color, fix font size ---
    for style in styles_root.iter(f'{W}style'):
        style_id = style.get(f'{W}styleId')
        if style_id == 'Heading5':
            rPr = style.find(f'{W}rPr')
            if rPr is not None:
                color = rPr.find(f'{W}color')
                if color is not None:
                    rPr.remove(color)
                sz = rPr.find(f'{W}sz')
                if sz is None:
                    sz = etree.SubElement(rPr, f'{W}sz')
                sz.set(f'{W}val', '24')
            print("  Fixed Heading 5")

    # --- Fix Table Grid style to 三线制 ---
    for style in styles_root.iter(f'{W}style'):
        style_id = style.get(f'{W}styleId')
        if style_id == 'TableGrid':
            tblPr = style.find(f'{W}tblPr')
            if tblPr is not None:
                borders = tblPr.find(f'{W}tblBorders')
                if borders is not None:
                    # 三线制: thick top/bottom only at table level, no sides, no inside
                    for b_name in ['left', 'right', 'insideH', 'insideV']:
                        b = borders.find(f'{W}{b_name}')
                        if b is not None:
                            borders.remove(b)
                    top = borders.find(f'{W}top')
                    if top is not None:
                        top.set(f'{W}sz', '12')   # 1.5pt
                    bottom = borders.find(f'{W}bottom')
                    if bottom is not None:
                        bottom.set(f'{W}sz', '12') # 1.5pt
            print("  Fixed Table Grid → 三线制 (3 lines)")

    styles_tree.write(styles_path, encoding='UTF-8', xml_declaration=True)

    # Repack
    fixed_ref = '/tmp/reference_fixed.docx'
    if os.path.exists(fixed_ref):
        os.remove(fixed_ref)
    with zipfile.ZipFile(fixed_ref, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(tmp_dir):
            for fname in files:
                full = os.path.join(root, fname)
                arc = os.path.relpath(full, tmp_dir)
                zout.write(full, arc)
    shutil.rmtree(tmp_dir)
    print(f"  Fixed reference saved: {fixed_ref}")
    return fixed_ref


# ── Step 2: Pandoc conversion ─────────────────────────────────
def run_pandoc(ref_docx):
    """Generate raw docx from markdown with fixed reference."""
    print("Step 2: Running pandoc...")
    raw = '/tmp/thesis_raw.docx'
    subprocess.run([
        'pandoc', MD_FILE, '-o', raw,
        '--reference-doc=' + ref_docx,
        '--from=markdown+smart-subscript', '--to=docx',
    ], check=True, cwd='/home/hyhml/my-notes')
    print(f"  Pandoc done: {raw} ({os.path.getsize(raw)} bytes)")
    return raw


# ── Step 3: Post-process docx OOXML ──────────────────────────
def post_process(raw_docx):
    """Apply all formatting fixes that pandoc cannot handle."""
    print("Step 3: Post-processing OOXML...")
    tmp_dir = '/tmp/thesis_unpacked'
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    with zipfile.ZipFile(raw_docx, 'r') as z:
        z.extractall(tmp_dir)

    doc_path = os.path.join(tmp_dir, 'word', 'document.xml')
    doc_tree = etree.parse(doc_path)
    doc_root = doc_tree.getroot()

    # 3a. Page setup
    fix_page_setup(doc_root)

    # 3b. Fix heading alignment (left), spacing, and jc issues
    fix_heading_alignment(doc_root)

    # 3c. Equation alignment (center + number right)
    fix_equation_alignment(doc_root)

    # 3d. Table captions (五号宋体, centered) — before text justification so jc=center skips indent
    fix_table_captions(doc_root)

    # 3d2. Figure captions (五号宋体, centered) — before text justification
    fix_figure_captions(doc_root)

    # 3e. Text justification (两端对齐) — body text only
    fix_text_justification(doc_root)

    # 3e1. Remove widow/orphan control from all paragraphs
    fix_widow_control(doc_root)

    # 3e2. References section (参考文献)
    fix_references(doc_root)

    # 3f. Table borders (apply 三线制)
    fix_table_borders(doc_root)

    # 3g. Table width (full page width)
    fix_table_width(doc_root)

    # 3g2. Table cell text alignment (center both ways)
    fix_table_cell_alignment(doc_root)

    # 3h. Math font + upright (no italic)
    fix_math_font_inplace(doc_root, tmp_dir)

    # 3i. Add header/footer
    add_header_footer(tmp_dir, doc_root)

    # Write document.xml
    # Use tostring to preserve all namespaces
    doc_new = etree.tostring(doc_root, encoding='UTF-8', xml_declaration=True, standalone=None)
    with open(doc_path, 'wb') as f:
        f.write(doc_new)

    # Repack
    if os.path.exists(OUTPUT):
        os.remove(OUTPUT)
    with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(tmp_dir):
            for fname in files:
                full = os.path.join(root, fname)
                arc = os.path.relpath(full, tmp_dir)
                zout.write(full, arc)
    shutil.rmtree(tmp_dir)
    print(f"  Done: {OUTPUT} ({os.path.getsize(OUTPUT)} bytes)")


# ── Heading alignment (第六条 表1) ─────────────────────────────
def fix_heading_alignment(doc_root):
    """Fix heading paragraphs: no indent, no bold, left-aligned, TNR numbers.
    WUST: 题序和标题之间空两个字符(2 en-spaces)，不加标点，左对齐.
    H1: 前面插入小四号空白行（空行有 jc=left + ind=0，不会被后续 justify 处理）
    """
    import re
    h1_count = 0
    h_count = 0
    for p in doc_root.iter(f'{W}p'):
        pPr = p.find(f'{W}pPr')
        if pPr is None:
            continue
        pStyle = pPr.find(f'{W}pStyle')
        if pStyle is None:
            continue
        style_val = pStyle.get(f'{W}val')
        if not style_val or 'Heading' not in style_val:
            continue

        # Override Normal style's firstLine — headings must be flush left (顶头)
        ind = pPr.find(f'{W}ind')
        if ind is None:
            ind = etree.Element(f'{W}ind')
            pPr.insert(0, ind)
        ind.set(f'{W}firstLine', '0')
        ind.set(f'{W}firstLineChars', '0')
        ind.set(f'{W}left', '0')
        try:
            del ind.attrib[f'{{{W_NS}}}hanging']
        except KeyError:
            pass

        # Fix jc: all headings left-aligned per WUST, except 致谢 centered
        heading_text = ''
        for t in p.iter(f'{W}t'):
            if t.text:
                heading_text += t.text
        is_thanks = heading_text.strip() == '致谢'

        jc = pPr.find(f'{W}jc')
        if jc is None:
            jc = etree.SubElement(pPr, f'{W}jc')
        jc.set(f'{W}val', 'center' if is_thanks else 'left')

        # H1 spacing: insert blank paragraph 小四号 before each H1 (skip 致谢)
        if style_val == 'Heading1' and not is_thanks:
            blank_p = etree.Element(f'{W}p')
            blank_pPr = etree.SubElement(blank_p, f'{W}pPr')
            blank_ind = etree.SubElement(blank_pPr, f'{W}ind')
            blank_ind.set(f'{W}firstLine', '0')
            blank_ind.set(f'{W}firstLineChars', '0')
            blank_spacing = etree.SubElement(blank_pPr, f'{W}spacing')
            blank_spacing.set(f'{W}line', '300')
            blank_spacing.set(f'{W}lineRule', 'auto')
            blank_jc = etree.SubElement(blank_pPr, f'{W}jc')
            blank_jc.set(f'{W}val', 'left')
            # Set font to 小四号 (12pt → sz=24) for correct line height
            blank_r = etree.SubElement(blank_p, f'{W}r')
            blank_rPr = etree.SubElement(blank_r, f'{W}rPr')
            blank_sz = etree.SubElement(blank_rPr, f'{W}sz')
            blank_sz.set(f'{W}val', '24')
            blank_szCs = etree.SubElement(blank_rPr, f'{W}szCs')
            blank_szCs.set(f'{W}val', '24')
            # Insert blank paragraph before H1
            parent = p.getparent()
            idx = list(parent).index(p)
            parent.insert(idx, blank_p)
            h1_count += 1

        # 致谢: insert page break before heading + blank line after (same as 参考文献)
        if is_thanks:
            parent = p.getparent()
            idx = list(parent).index(p)
            # Page break paragraph before 致谢
            pb_p = etree.Element(f'{W}p')
            pb_pPr = etree.SubElement(pb_p, f'{W}pPr')
            pb_r = etree.SubElement(pb_p, f'{W}r')
            pb_br = etree.SubElement(pb_r, f'{W}br')
            pb_br.set(f'{W}type', 'page')
            parent.insert(idx, pb_p)
            # Blank line paragraph after 致谢 heading
            thanks_blank = etree.Element(f'{W}p')
            thanks_blank_pPr = etree.SubElement(thanks_blank, f'{W}pPr')
            thanks_blank_ind = etree.SubElement(thanks_blank_pPr, f'{W}ind')
            thanks_blank_ind.set(f'{W}firstLine', '0')
            thanks_blank_ind.set(f'{W}firstLineChars', '0')
            thanks_blank_spacing = etree.SubElement(thanks_blank_pPr, f'{W}spacing')
            thanks_blank_spacing.set(f'{W}line', '300')
            thanks_blank_spacing.set(f'{W}lineRule', 'auto')
            thanks_blank_jc = etree.SubElement(thanks_blank_pPr, f'{W}jc')
            thanks_blank_jc.set(f'{W}val', 'left')
            thanks_blank_r = etree.SubElement(thanks_blank, f'{W}r')
            thanks_blank_rPr = etree.SubElement(thanks_blank_r, f'{W}rPr')
            thanks_blank_sz = etree.SubElement(thanks_blank_rPr, f'{W}sz')
            thanks_blank_sz.set(f'{W}val', '24')
            thanks_blank_szCs = etree.SubElement(thanks_blank_rPr, f'{W}szCs')
            thanks_blank_szCs.set(f'{W}val', '24')
            # Insert after 致谢 heading (index shifted by page break insertion)
            parent.insert(idx + 2, thanks_blank)
            # Set 致谢 heading font size to 三号 (sz=32)
            for r in p.iter(f'{W}r'):
                rPr = r.find(f'{W}rPr')
                if rPr is None:
                    rPr = etree.SubElement(r, f'{W}rPr')
                for sz_tag in ['sz', 'szCs']:
                    sz_el = rPr.find(f'{W}{sz_tag}')
                    if sz_el is None:
                        sz_el = etree.SubElement(rPr, f'{W}{sz_tag}')
                    sz_el.set(f'{W}val', '32')

        # Ensure all heading runs have explicit rFonts and no bold:
        # TNR for western/numbers, 黑体 for CJK, no bold
        for r in p.iter(f'{W}r'):
            rPr = r.find(f'{W}rPr')
            if rPr is None:
                rPr = etree.Element(f'{W}rPr')
                r.insert(0, rPr)
            # Remove bold from run level
            for b_tag in ['b', 'bCs']:
                b_el = rPr.find(f'{W}{b_tag}')
                if b_el is not None:
                    rPr.remove(b_el)
            rFonts = rPr.find(f'{W}rFonts')
            if rFonts is None:
                rFonts = etree.SubElement(rPr, f'{W}rFonts')
            # Strip theme references
            for attr in ['asciiTheme', 'eastAsiaTheme', 'hAnsiTheme', 'cstheme']:
                try:
                    del rFonts.attrib[f'{{{W_NS}}}{attr}']
                except KeyError:
                    pass
            rFonts.set(f'{W}ascii', 'Times New Roman')
            rFonts.set(f'{W}hAnsi', 'Times New Roman')
            rFonts.set(f'{W}eastAsia', '黑体')
            rFonts.set(f'{W}cs', 'Times New Roman')

        # Two spaces between number and title: "3.1 X" → "3.1  X"
        # (题序和标题之间空两个字符)
        for t in p.iter(f'{W}t'):
            if t.text:
                t.text = re.sub(r'^(\d+(?:\.\d+)*) ', r'\1  ', t.text)

        h_count += 1
    print(f"  Heading alignment: {h_count} headings → left-aligned, no bold, 2-en-space separator ({h1_count} H1 with blank line) [致谢 centered]")

# ── Equation alignment (第十条-四) ──────────────────────────────
def fix_equation_alignment(doc_root):
    """Center equations and right-align their numbers using Word native eqArr.

    Pandoc converts $$ ... \\qquad (X.Y) $$ to OMML with the number and
    \\qquad space either fused into the last m:r or split across separate
    m:r elements (\\u2001\\u2001, (, 3.1, )).  We transform this into
    Word's native eqArr + # syntax, which Word renders as centered
    equation with right-aligned number automatically.
    """
    import re
    eq_count = 0
    for p in doc_root.iter(f'{W}p'):
        omp = p.find(f'.//{M}oMathPara')
        if omp is None:
            continue

        oMaths = omp.findall(f'{M}oMath')
        if not oMaths:
            continue

        last_oMath = oMaths[-1]

        # ── Check for matrix structure (from pandoc \begin{aligned}) ──
        m_matrix = last_oMath.find(f'{M}m')
        if m_matrix is not None:
            # Matrix equation: runs are nested in m:m → m:mr → m:e
            all_mr = list(m_matrix.iter(f'{M}r'))
            if not all_mr:
                continue
        else:
            # Regular equation: runs are direct children of oMath
            all_mr = last_oMath.findall(f'{M}r')
            if not all_mr:
                continue

        # ── Collect run data ──
        run_data = []  # [(m_r, full_text)]
        for mr in all_mr:
            texts = []
            for mt in mr.iter(f'{M}t'):
                if mt.text:
                    texts.append(mt.text)
            run_data.append((mr, ''.join(texts)))

        # ── Find number pattern at end, across possibly multiple runs ──
        # Concatenate the last few runs' text to find the pattern
        all_runs_text = ''.join(t for _, t in run_data)
        m_end = re.search(r'[ \s]*\((\d+\.\d+(?:\.\d+)?)\)\s*$', all_runs_text)
        if not m_end:
            continue

        num_val = m_end.group(1)      # e.g. '3.1'
        num_start = m_end.start()     # position in concatenated text

        # ── Locate the run and m:t where the number pattern starts ──
        cumulative = 0
        split_mr_idx = None   # first run that contains part of the number pattern
        split_mt = None       # the specific m:t element to truncate
        offset_in_mt = None   # where to cut in that m:t's text

        for idx, (mr, text) in enumerate(run_data):
            if cumulative + len(text) > num_start:
                split_mr_idx = idx
                offset_in_mt = num_start - cumulative
                # Find the specific m:t element
                mt_cum = 0
                for mt in mr.iter(f'{M}t'):
                    if mt.text:
                        if mt_cum + len(mt.text) > offset_in_mt:
                            split_mt = mt
                            offset_in_mt -= mt_cum
                            break
                        mt_cum += len(mt.text)
                break
            cumulative += len(text)

        if split_mt is None:
            continue

        # ── Remove runs that are entirely part of the number/spacer ──
        # (runs after split_mr that contain only number/spacer fragments)
        runs_to_remove = []
        for idx in range(len(run_data) - 1, split_mr_idx, -1):
            t = run_data[idx][1].strip()
            if t == '' or t == '(' or t == ')' or re.match(r'^\d+\.\d+(?:\.\d+)?$', t):
                # Also check if it's purely spaces
                if run_data[idx][1].strip() == '':
                    pass  # pure whitespace run
                runs_to_remove.append(run_data[idx][0])
            else:
                # This run has non-number content — shouldn't happen if regex matched
                break

        for mr in reversed(runs_to_remove):
            mr.getparent().remove(mr)

        # ── Truncate the split run; append '#' for regular eqs (matrix eqs add it at eqArr level) ──
        if m_matrix is not None:
            split_mt.text = split_mt.text[:offset_in_mt].rstrip()
        else:
            split_mt.text = split_mt.text[:offset_in_mt].rstrip() + '#'

        # ── Build m:d (delimiter) for the number ──
        m_d = etree.Element(f'{M}d')
        m_dPr = etree.SubElement(m_d, f'{M}dPr')
        m_dCtrlPr = etree.SubElement(m_dPr, f'{M}ctrlPr')
        m_dCtrl_rPr = etree.SubElement(m_dCtrlPr, f'{W}rPr')
        m_dCtrl_fonts = etree.SubElement(m_dCtrl_rPr, f'{W}rFonts')
        m_dCtrl_fonts.set(f'{W}ascii', 'Times New Roman')
        m_dCtrl_fonts.set(f'{W}hAnsi', 'Times New Roman')
        m_d_e = etree.SubElement(m_d, f'{M}e')
        m_d_mr = etree.SubElement(m_d_e, f'{M}r')
        m_d_mrPr = etree.SubElement(m_d_mr, f'{M}rPr')
        m_d_mrFonts = etree.SubElement(m_d_mrPr, f'{W}rFonts')
        m_d_mrFonts.set(f'{W}ascii', 'Times New Roman')
        m_d_mrFonts.set(f'{W}hAnsi', 'Times New Roman')
        m_d_mrSty = etree.SubElement(m_d_mrPr, f'{M}sty')
        m_d_mrSty.set(f'{M}val', 'p')   # equation numbers upright
        m_d_mt = etree.SubElement(m_d_mr, f'{M}t')
        m_d_mt.text = num_val

        # ── Build eqArr with number ──
        eqArr = etree.Element(f'{M}eqArr')
        eqArrPr = etree.SubElement(eqArr, f'{M}eqArrPr')
        maxDist = etree.SubElement(eqArrPr, f'{M}maxDist')
        maxDist.set(f'{M}val', '1')
        eqCtrlPr = etree.SubElement(eqArrPr, f'{M}ctrlPr')
        eqCtrl_rPr = etree.SubElement(eqCtrlPr, f'{W}rPr')
        eqCtrl_fonts = etree.SubElement(eqCtrl_rPr, f'{W}rFonts')
        eqCtrl_fonts.set(f'{W}ascii', 'Times New Roman')
        eqCtrl_fonts.set(f'{W}hAnsi', 'Times New Roman')

        if m_matrix is not None:
            # ── Inner eqArr: multi-line equation with & alignment ──
            # No maxDist so columns spread to fill available width, matching
            # the horizontal span of regular equations.
            inner_eqArr = etree.Element(f'{M}eqArr')
            inner_eqArrPr = etree.SubElement(inner_eqArr, f'{M}eqArrPr')

            m_mrs = m_matrix.findall(f'{M}mr')

            for row_idx, mr_row in enumerate(m_mrs):
                m_elem_row = etree.SubElement(inner_eqArr, f'{M}e')
                cells = mr_row.findall(f'{M}e')

                # Column 0 (right-aligned part; may be empty)
                for child in list(cells[0]) if cells else []:
                    cells[0].remove(child)
                    m_elem_row.append(child)

                # & alignment marker
                amp_r = etree.SubElement(m_elem_row, f'{M}r')
                amp_rPr = etree.SubElement(amp_r, f'{M}rPr')
                amp_sty = etree.SubElement(amp_rPr, f'{M}sty')
                amp_sty.set(f'{M}val', 'p')
                amp_mt = etree.SubElement(amp_r, f'{M}t')
                amp_mt.text = '&'

                # Column 1+ (left-aligned part)
                # Strip leading whitespace (e.g. \quad →  ) so content
                # aligns at & rather than being offset by LaTeX spacing.
                first_col1_child = True
                for cell in cells[1:]:
                    for child in list(cell):
                        cell.remove(child)
                        if first_col1_child:
                            for mt in child.iter(f'{M}t'):
                                if mt.text:
                                    mt.text = mt.text.lstrip('      ')
                                    break
                            first_col1_child = False
                        m_elem_row.append(child)

            last_oMath.remove(m_matrix)

            # ── Outer eqArr wraps inner + # + number ──
            # The outer eqArr uses the same maxDist=1 as regular equations,
            # so the number position is consistent.  The number is vertically
            # centred relative to the equation block because it is a sibling
            # of the inner eqArr, not attached to a single row.
            m_elem = etree.SubElement(eqArr, f'{M}e')
            m_elem.append(inner_eqArr)
            hash_r = etree.SubElement(m_elem, f'{M}r')
            hash_mt = etree.SubElement(hash_r, f'{M}t')
            hash_mt.text = '#'
            m_elem.append(m_d)
        else:
            # Regular equation: insert m:d after split run, wrap all children
            target_mr = run_data[split_mr_idx][0]
            mr_idx = list(last_oMath).index(target_mr)
            last_oMath.insert(mr_idx + 1, m_d)

            children = list(last_oMath)
            m_elem = etree.SubElement(eqArr, f'{M}e')
            for child in children:
                last_oMath.remove(child)
                m_elem.append(child)

        last_oMath.append(eqArr)

        # Remove oMathParaPr/jc — eqArr handles alignment natively
        ompPr = omp.find(f'{M}oMathParaPr')
        if ompPr is not None:
            m_jc = ompPr.find(f'{M}jc')
            if m_jc is not None:
                ompPr.remove(m_jc)

        eq_count += 1
    print(f"  Equation alignment: {eq_count} equations → eqArr (Word native)")


# ── Page setup (第十二条) ─────────────────────────────────────
def fix_page_setup(doc_root):
    """A4 paper, 2.5cm margins, correct header/footer distances."""
    for sectPr in doc_root.iter(f'{W}sectPr'):
        pgSz = sectPr.find(f'{W}pgSz')
        if pgSz is not None:
            pgSz.set(f'{W}w', '11906')   # A4 width (21cm)
            pgSz.set(f'{W}h', '16838')   # A4 height (29.7cm)

        pgMar = sectPr.find(f'{W}pgMar')
        if pgMar is not None:
            pgMar.set(f'{W}top', '1417')     # 2.5cm
            pgMar.set(f'{W}bottom', '1417')  # 2.5cm
            pgMar.set(f'{W}left', '1417')    # 2.5cm
            pgMar.set(f'{W}right', '1417')   # 2.5cm
            pgMar.set(f'{W}header', '1134')  # 2cm (header distance from top edge)
            pgMar.set(f'{W}footer', '992')   # 1.75cm (footer distance from bottom edge)
            pgMar.set(f'{W}gutter', '0')
        print("  Page setup: A4, margins 2.5cm, header=2cm, footer=1.75cm")
        break  # Only fix first section


# ── Text justification (两端对齐) ────────────────────────────
def fix_text_justification(doc_root):
    """Set body paragraphs to 两端对齐 (jc=both) with 2-char first-line indent.
    Skip headings, captions, equation paragraphs, and specially-formatted ones."""
    count = 0
    for p in doc_root.iter(f'{W}p'):
        pPr = p.find(f'{W}pPr')
        if pPr is None:
            continue
        # Skip headings (outlineLvl or pStyle=Heading*)
        if pPr.find(f'{W}outlineLvl') is not None:
            continue
        pStyle = pPr.find(f'{W}pStyle')
        if pStyle is not None:
            style_val = pStyle.get(f'{W}val')
            if style_val and 'Heading' in style_val:
                continue
        # Skip paragraphs containing display math (handled by fix_equation_alignment)
        if p.find(f'.//{M}oMathPara') is not None:
            continue
        # Skip table cell paragraphs (handled by fix_table_cell_alignment)
        # — check if parent chain includes a w:tc
        parent = p.getparent()
        is_in_table = False
        while parent is not None:
            if parent.tag == f'{W}tc':
                is_in_table = True
                break
            parent = parent.getparent()
        if is_in_table:
            continue

        # Image paragraphs: no indent, centered, kept with caption
        if p.find(f'.//{W}drawing') is not None:
            # Explicitly override Normal style's indent
            ind = pPr.find(f'{W}ind')
            if ind is None:
                ind = etree.SubElement(pPr, f'{W}ind')
            ind.set(f'{W}firstLine', '0')
            ind.set(f'{W}firstLineChars', '0')
            # Center alignment
            jc_img = pPr.find(f'{W}jc')
            if jc_img is None:
                jc_img = etree.SubElement(pPr, f'{W}jc')
            jc_img.set(f'{W}val', 'center')
            # Keep with next (caption) and keep lines together
            if pPr.find(f'{W}keepNext') is None:
                pPr.insert(0, etree.Element(f'{W}keepNext'))
            if pPr.find(f'{W}keepLines') is None:
                pPr.insert(0, etree.Element(f'{W}keepLines'))
            count += 1
            continue

        # Skip if already has jc set (centered captions, etc.)
        jc = pPr.find(f'{W}jc')
        if jc is not None:
            continue

        # Set first-line indent: 2 Chinese characters
        ind = pPr.find(f'{W}ind')
        if ind is None:
            ind = etree.Element(f'{W}ind')
            pPr.insert(0, ind)
        ind.set(f'{W}firstLineChars', '200')

        # Add justification (两端对齐)
        jc_new = etree.Element(f'{W}jc')
        jc_new.set(f'{W}val', 'both')
        # Insert after spacing or ind
        spacing = pPr.find(f'{W}spacing')
        if ind is not None:
            ins_idx = list(pPr).index(ind) + 1
            pPr.insert(ins_idx, jc_new)
        elif spacing is not None:
            ins_idx = list(pPr).index(spacing) + 1
            pPr.insert(ins_idx, jc_new)
        else:
            pPr.insert(0, jc_new)
        count += 1
    print(f"  Text justification: {count} paragraphs → 两端对齐 + 2-char indent")


# ── Widow/Orphan control ────────────────────────────────────
def fix_widow_control(doc_root):
    """Disable widow/orphan control for all body paragraphs."""
    count = 0
    for p in doc_root.iter(f'{W}p'):
        pPr = p.find(f'{W}pPr')
        if pPr is None:
            pPr = etree.Element(f'{W}pPr')
            p.insert(0, pPr)
        wc = pPr.find(f'{W}widowControl')
        if wc is not None:
            wc.set(f'{W}val', 'false')
        else:
            etree.SubElement(pPr, f'{W}widowControl', {f'{W}val': 'false'})
        count += 1
    print(f"  Widow control: disabled in {count} paragraphs")


# ── References section (第十条-六) ──────────────────────────────
def fix_references(doc_root):
    """Fix 参考文献 section:
    - Title: 三号黑体居中
    - Blank line after title
    - Reference items: flush-left (no first-line indent)
    """
    body = doc_root.find(f'{W}body')
    if body is None:
        return

    # Find the 参考文献 heading paragraph
    ref_heading = None
    for p in body.iter(f'{W}p'):
        pPr = p.find(f'{W}pPr')
        if pPr is None:
            continue
        pStyle = pPr.find(f'{W}pStyle')
        if pStyle is None:
            continue
        style_val = pStyle.get(f'{W}val')
        if style_val != 'Heading1':
            continue
        # Collect all run texts
        texts = []
        for r in p.iter(f'{W}r'):
            t = r.find(f'{W}t')
            if t is not None and t.text:
                texts.append(t.text)
        full_text = ''.join(texts)
        if '参考文献' in full_text:
            ref_heading = p
            break

    if ref_heading is None:
        print("  References: 参考文献 heading not found — skip")
        return

    # Fix heading: center, 三号 (sz=32)
    pPr = ref_heading.find(f'{W}pPr')
    jc = pPr.find(f'{W}jc')
    if jc is None:
        jc = etree.SubElement(pPr, f'{W}jc')
    jc.set(f'{W}val', 'center')
    for r in ref_heading.iter(f'{W}r'):
        rPr = r.find(f'{W}rPr')
        if rPr is None:
            rPr = etree.SubElement(r, f'{W}rPr')
        for sz_tag in ['sz', 'szCs']:
            sz = rPr.find(f'{W}{sz_tag}')
            if sz is None:
                sz = etree.SubElement(rPr, f'{W}{sz_tag}')
            sz.set(f'{W}val', '32')

    # Remove blank paragraph before 参考文献 (inserted by H1 processing)
    parent = ref_heading.getparent()
    idx = list(parent).index(ref_heading)
    if idx > 0:
        prev_p = list(parent)[idx - 1]
        prev_pPr = prev_p.find(f'{W}pPr')
        if prev_pPr is not None:
            prev_style = prev_pPr.find(f'{W}pStyle')
            if prev_style is None:  # blank para has no style
                has_text = bool(prev_p.find(f'.//{W}t'))
                if not has_text:
                    parent.remove(prev_p)

    # Insert page break paragraph right before the heading (after blank removal)
    idx = list(parent).index(ref_heading)
    page_break_p = etree.Element(f'{W}p')
    page_break_pPr = etree.SubElement(page_break_p, f'{W}pPr')
    page_break_r = etree.SubElement(page_break_p, f'{W}r')
    page_break_br = etree.SubElement(page_break_r, f'{W}br')
    page_break_br.set(f'{W}type', 'page')
    parent.insert(idx, page_break_p)

    # Insert blank line after heading
    blank_p = etree.Element(f'{W}p')
    blank_pPr = etree.SubElement(blank_p, f'{W}pPr')
    blank_ind = etree.SubElement(blank_pPr, f'{W}ind')
    blank_ind.set(f'{W}firstLine', '0')
    blank_ind.set(f'{W}firstLineChars', '0')
    blank_spacing = etree.SubElement(blank_pPr, f'{W}spacing')
    blank_spacing.set(f'{W}line', '300')
    blank_spacing.set(f'{W}lineRule', 'auto')
    blank_jc = etree.SubElement(blank_pPr, f'{W}jc')
    blank_jc.set(f'{W}val', 'left')
    blank_r = etree.SubElement(blank_p, f'{W}r')
    blank_rPr = etree.SubElement(blank_r, f'{W}rPr')
    blank_sz = etree.SubElement(blank_rPr, f'{W}sz')
    blank_sz.set(f'{W}val', '24')
    blank_szCs = etree.SubElement(blank_rPr, f'{W}szCs')
    blank_szCs.set(f'{W}val', '24')
    # Insert after heading
    idx = list(parent).index(ref_heading)
    parent.insert(idx + 1, blank_p)

    # Fix all reference paragraphs: left-align, no indent, fix NBSPs
    ref_count = 0
    nbsp_fixed = 0
    found_ref_heading = False
    for p in body.iter(f'{W}p'):
        pPr = p.find(f'{W}pPr')
        if pPr is None:
            continue
        if p == ref_heading:
            found_ref_heading = True
            continue
        if not found_ref_heading:
            continue
        # Stop at next heading (致谢)
        pStyle = pPr.find(f'{W}pStyle')
        if pStyle is not None:
            style_val = pStyle.get(f'{W}val')
            if style_val and 'Heading' in style_val:
                break

        # Remove first-line indent
        ind = pPr.find(f'{W}ind')
        if ind is not None:
            ind.set(f'{W}firstLine', '0')
            ind.set(f'{W}firstLineChars', '0')

        # Fix jc: references should be both (两端对齐) for clean appearance
        jc = pPr.find(f'{W}jc')
        if jc is None:
            jc = etree.SubElement(pPr, f'{W}jc')
        jc.set(f'{W}val', 'both')

        # Fix text runs: replace NBSPs, restore double-space after bracket
        for t in p.iter(f'{W}t'):
            if t.text:
                # Replace pandoc smart-typography NBSPs (et al.→title)
                if '\xa0' in t.text:
                    t.text = t.text.replace('\xa0', ' ')
                    nbsp_fixed += 1
                # Restore two spaces after bracket: pandoc collapses "[1]  X" → "[1] X"
                import re
                t.text = re.sub(r'^\[(\d+)\] ', r'[\1]  ', t.text)

        ref_count += 1

    print(f"  References: title centered, {ref_count} items → flush-left, {nbsp_fixed} NBSPs fixed")


# ── Table captions (第十条-二) ────────────────────────────────
def fix_table_captions(doc_root):
    """Fix table captions to 五号宋体(10.5pt) centered above table.
    'Table X-Y: ...' → '表X.Y ...'
    """
    body = doc_root.find(f'{W}body')
    if body is None:
        return
    count = 0
    elements = list(body)
    for i, el in enumerate(elements):
        # Check if next element is a table
        if i + 1 >= len(elements):
            continue
        next_el = elements[i + 1]
        if next_el.tag != f'{W}tbl':
            continue
        # Check if current paragraph contains "Table" or "表" in text
        texts = []
        for t in el.iter(f'{W}t'):
            if t.text:
                texts.append(t.text)
        full_text = ''.join(texts)
        if not full_text.strip():
            continue
        if not ('Table' in full_text or '表' in full_text):
            continue

        # This is a table caption paragraph
        pPr = el.find(f'{W}pPr')
        if pPr is None:
            pPr = etree.Element(f'{W}pPr')
            el.insert(0, pPr)

        # Set centered alignment
        jc = pPr.find(f'{W}jc')
        if jc is None:
            jc = etree.SubElement(pPr, f'{W}jc')
        jc.set(f'{W}val', 'center')

        # Keep caption with table body on same page
        if pPr.find(f'{W}keepNext') is None:
            pPr.insert(0, etree.Element(f'{W}keepNext'))
        if pPr.find(f'{W}keepLines') is None:
            pPr.insert(0, etree.Element(f'{W}keepLines'))

        # Remove any first-line indent
        ind = pPr.find(f'{W}ind')
        if ind is None:
            ind = etree.SubElement(pPr, f'{W}ind')
        ind.set(f'{W}firstLine', '0')
        ind.set(f'{W}firstLineChars', '0')

        # Fix paragraph style (remove BodyText, use Normal)
        pStyle = pPr.find(f'{W}pStyle')
        if pStyle is not None and pStyle.get(f'{W}val') == 'BodyText':
            pPr.remove(pStyle)

        # Fix runs: set font to 宋体 五号(sz=21)
        for r in el.iter(f'{W}r'):
            rPr = r.find(f'{W}rPr')
            if rPr is None:
                rPr = etree.Element(f'{W}rPr')
                r.insert(0, rPr)
            rFonts = rPr.find(f'{W}rFonts')
            if rFonts is None:
                rFonts = etree.SubElement(rPr, f'{W}rFonts')
            rFonts.set(f'{W}ascii', 'Times New Roman')
            rFonts.set(f'{W}hAnsi', 'Times New Roman')
            rFonts.set(f'{W}eastAsia', '宋体')
            sz = rPr.find(f'{W}sz')
            if sz is None:
                sz = etree.SubElement(rPr, f'{W}sz')
            sz.set(f'{W}val', '21')  # 五号 = 10.5pt
            szCs = rPr.find(f'{W}szCs')
            if szCs is None:
                szCs = etree.SubElement(rPr, f'{W}szCs')
            szCs.set(f'{W}val', '21')

        # Fix "Table X-Y" or "表X-Y" → "表X.Y"
        for t in el.iter(f'{W}t'):
            if t.text:
                import re
                t.text = re.sub(r'[Tt]able\s+(\d+)-(\d+)', r'表\1.\2', t.text)
                t.text = re.sub(r'表\s*(\d+)-(\d+)', r'表\1.\2', t.text)

        count += 1
    print(f"  Table captions: {count} fixed (五号宋体, centered)")


# ── Figure captions (第十条-二) ────────────────────────────────
def fix_figure_captions(doc_root):
    """Fix figure captions to 五号宋体(10.5pt) centered below image.
    Detects paragraphs with '图' text that follow an image (w:drawing).
    """
    body = doc_root.find(f'{W}body')
    if body is None:
        return
    count = 0
    elements = list(body)
    for i, el in enumerate(elements):
        if i == 0:
            continue
        if el.tag != f'{W}p':
            continue
        # Check if previous element is a paragraph containing a drawing
        prev_el = elements[i - 1]
        if prev_el.tag != f'{W}p':
            continue
        if prev_el.find(f'.//{W}drawing') is None:
            continue
        # Check if current paragraph contains '图' in text
        texts = []
        for t in el.iter(f'{W}t'):
            if t.text:
                texts.append(t.text)
        full_text = ''.join(texts)
        if '图' not in full_text:
            continue

        # This is a figure caption paragraph
        pPr = el.find(f'{W}pPr')
        if pPr is None:
            pPr = etree.Element(f'{W}pPr')
            el.insert(0, pPr)

        # Set centered alignment
        jc = pPr.find(f'{W}jc')
        if jc is None:
            jc = etree.SubElement(pPr, f'{W}jc')
        jc.set(f'{W}val', 'center')

        # Keep with previous (image) and keep lines together
        if pPr.find(f'{W}keepNext') is None:
            pPr.insert(0, etree.Element(f'{W}keepNext'))
        if pPr.find(f'{W}keepLines') is None:
            pPr.insert(0, etree.Element(f'{W}keepLines'))

        # Fix paragraph style (remove BodyText, use Normal)
        pStyle = pPr.find(f'{W}pStyle')
        if pStyle is not None and pStyle.get(f'{W}val') == 'BodyText':
            pPr.remove(pStyle)

        # Remove any first-line indent from caption
        ind = pPr.find(f'{W}ind')
        if ind is None:
            ind = etree.SubElement(pPr, f'{W}ind')
        ind.set(f'{W}firstLine', '0')
        ind.set(f'{W}firstLineChars', '0')

        # Fix runs: set font to 宋体 五号(sz=21)
        for r in el.iter(f'{W}r'):
            rPr = r.find(f'{W}rPr')
            if rPr is None:
                rPr = etree.Element(f'{W}rPr')
                r.insert(0, rPr)
            rFonts = rPr.find(f'{W}rFonts')
            if rFonts is None:
                rFonts = etree.SubElement(rPr, f'{W}rFonts')
            rFonts.set(f'{W}ascii', 'Times New Roman')
            rFonts.set(f'{W}hAnsi', 'Times New Roman')
            rFonts.set(f'{W}eastAsia', '宋体')
            sz = rPr.find(f'{W}sz')
            if sz is None:
                sz = etree.SubElement(rPr, f'{W}sz')
            sz.set(f'{W}val', '21')  # 五号 = 10.5pt
            szCs = rPr.find(f'{W}szCs')
            if szCs is None:
                szCs = etree.SubElement(rPr, f'{W}szCs')
            szCs.set(f'{W}val', '21')

        count += 1
    print(f"  Figure captions: {count} fixed (五号宋体, centered)")


# ── Table borders: 三线制 (第十条-二) ─────────────────────────
def fix_table_borders(doc_root):
    """Apply 三线制 to all tables:
    - Top/bottom: 1.5pt thick (sz=12) at table level
    - Header row bottom: 0.75pt thin (sz=6) on first row cells
    - No left/right/inside borders
    """
    count = 0
    for tbl in doc_root.iter(f'{W}tbl'):
        tblPr = tbl.find(f'{W}tblPr')
        if tblPr is None:
            tblPr = etree.Element(f'{W}tblPr')
            tbl.insert(0, tblPr)

        # Remove existing borders
        old_borders = tblPr.find(f'{W}tblBorders')
        if old_borders is not None:
            tblPr.remove(old_borders)

        # Table-level: only top and bottom thick lines
        borders = etree.SubElement(tblPr, f'{W}tblBorders')

        top = etree.SubElement(borders, f'{W}top')
        top.set(f'{W}val', 'single')
        top.set(f'{W}sz', '12')       # 1.5pt
        top.set(f'{W}space', '0')
        top.set(f'{W}color', '000000')

        bottom = etree.SubElement(borders, f'{W}bottom')
        bottom.set(f'{W}val', 'single')
        bottom.set(f'{W}sz', '12')    # 1.5pt
        bottom.set(f'{W}space', '0')
        bottom.set(f'{W}color', '000000')

        # No left, right, insideH, insideV — only 2 lines at table level

        # Header row: add thin bottom border to each cell (0.75pt = sz 6)
        trs = tbl.findall(f'{W}tr')
        if trs:
            header_row = trs[0]
            for tc in header_row.findall(f'{W}tc'):
                tcPr = tc.find(f'{W}tcPr')
                if tcPr is None:
                    tcPr = etree.Element(f'{W}tcPr')
                    tc.insert(0, tcPr)
                # Remove existing cell borders
                old_tc_borders = tcPr.find(f'{W}tcBorders')
                if old_tc_borders is not None:
                    tcPr.remove(old_tc_borders)
                tc_borders = etree.SubElement(tcPr, f'{W}tcBorders')
                tc_bottom = etree.SubElement(tc_borders, f'{W}bottom')
                tc_bottom.set(f'{W}val', 'single')
                tc_bottom.set(f'{W}sz', '6')    # 0.75pt
                tc_bottom.set(f'{W}space', '0')
                tc_bottom.set(f'{W}color', '000000')

            # Keep all rows together on same page
            for j, tr in enumerate(trs):
                trPr = tr.find(f'{W}trPr')
                if trPr is None:
                    trPr = etree.Element(f'{W}trPr')
                    tr.insert(0, trPr)
                # keepLines: don't split row across pages
                if trPr.find(f'{W}keepLines') is None:
                    trPr.insert(0, etree.Element(f'{W}keepLines'))
                # keepNext: stay with next row or note (all rows including last)
                if trPr.find(f'{W}keepNext') is None:
                    trPr.insert(0, etree.Element(f'{W}keepNext'))

        count += 1
    print(f"  Table borders: {count} tables → 三线制 (3 lines only)")


# ── Table width (第十条-二) ────────────────────────────────────
def fix_table_width(doc_root):
    """Set all tables to full page width (100% = 5000 pct)."""
    count = 0
    for tbl in doc_root.iter(f'{W}tbl'):
        tblPr = tbl.find(f'{W}tblPr')
        if tblPr is None:
            tblPr = etree.Element(f'{W}tblPr')
            tbl.insert(0, tblPr)

        tblW = tblPr.find(f'{W}tblW')
        if tblW is None:
            tblW = etree.SubElement(tblPr, f'{W}tblW')
        tblW.set(f'{W}w', '5000')
        tblW.set(f'{W}type', 'pct')
        count += 1
    print(f"  Table width: {count} tables → 100% page width")


# ── Table cell alignment ──────────────────────────────────────
def fix_table_cell_alignment(doc_root):
    """Format all table cell text: 五号字, centered horizontally and vertically."""
    tbl_count = 0
    cell_count = 0
    for tbl in doc_root.iter(f'{W}tbl'):
        for tc in tbl.iter(f'{W}tc'):
            tcPr = tc.find(f'{W}tcPr')
            if tcPr is None:
                tcPr = etree.Element(f'{W}tcPr')
                tc.insert(0, tcPr)
            # Vertical alignment
            vAlign = tcPr.find(f'{W}vAlign')
            if vAlign is None:
                vAlign = etree.SubElement(tcPr, f'{W}vAlign')
            vAlign.set(f'{W}val', 'center')

            # Horizontal alignment and font size for every paragraph in this cell
            for p in tc.iter(f'{W}p'):
                pPr = p.find(f'{W}pPr')
                if pPr is None:
                    pPr = etree.Element(f'{W}pPr')
                    p.insert(0, pPr)
                # Explicitly zero first-line indent — table text must not have
                # leading spaces. Must set explicitly, not remove, because
                # Normal style inherits firstLineChars=200.
                ind = pPr.find(f'{W}ind')
                if ind is None:
                    ind = etree.Element(f'{W}ind')
                    pPr.insert(0, ind)
                ind.set(f'{W}firstLine', '0')
                ind.set(f'{W}firstLineChars', '0')
                jc = pPr.find(f'{W}jc')
                if jc is None:
                    jc = etree.SubElement(pPr, f'{W}jc')
                jc.set(f'{W}val', 'center')

                # Set font: 五号 (sz=21) for every run in table cells
                for r in p.iter(f'{W}r'):
                    rPr = r.find(f'{W}rPr')
                    if rPr is None:
                        rPr = etree.Element(f'{W}rPr')
                        r.insert(0, rPr)
                    sz = rPr.find(f'{W}sz')
                    if sz is None:
                        sz = etree.SubElement(rPr, f'{W}sz')
                    sz.set(f'{W}val', '21')     # 五号 = 10.5pt
                    szCs = rPr.find(f'{W}szCs')
                    if szCs is None:
                        szCs = etree.SubElement(rPr, f'{W}szCs')
                    szCs.set(f'{W}val', '21')

                cell_count += 1
        tbl_count += 1
    print(f"  Table cell alignment: {tbl_count} tables, {cell_count} cells → 五号, centered")

# ── Math font (第十条-四) ─────────────────────────────────────
def fix_math_font_inplace(doc_root, tmp_dir):
    """Set math font to Times New Roman, with intelligent italic/upright.

    ISO 80000-2 rules:
    - Single-letter variables → italic
    - Multi-letter identifiers, function names, descriptive text → upright
    - Constants (π) and operators (Δ) → upright
    - Numbers → upright

    Pandoc often splits multi-letter text across adjacent m:r elements,
    so we classify by concatenating sibling runs under the same parent.
    """
    import re
    # Known math constants and operators (always upright per ISO 80000-2)
    UPRIGHT_CONSTANTS = {'π', 'Δ'}
    # Math parent elements to consider for sibling context
    MATH_PARENTS = {f'{M}sub', f'{M}sup', f'{M}e', f'{M}num', f'{M}den',
                    f'{M}deg', f'{M}lim', f'{M}oMath'}

    total = 0
    italic = 0
    upright = 0
    for m_r in doc_root.iter(f'{M}r'):
        total += 1
        m_rPr = m_r.find(f'{M}rPr')
        if m_rPr is None:
            m_rPr = etree.Element(f'{M}rPr')
            m_r.insert(0, m_rPr)
        # Always set TNR font
        rFonts = m_rPr.find(f'{W}rFonts')
        if rFonts is None:
            rFonts = etree.SubElement(m_rPr, f'{W}rFonts')
        rFonts.set(f'{W}ascii', 'Times New Roman')
        rFonts.set(f'{W}hAnsi', 'Times New Roman')
        rFonts.set(f'{W}cs', 'Times New Roman')

        # Get this run's text
        own_text = ''
        for mt in m_r.iter(f'{M}t'):
            if mt.text:
                own_text += mt.text
        own_text = own_text.strip()

        if not own_text:
            italic += 1
            continue

        # ── Get context: concatenate all sibling m:r text under same parent ──
        parent = m_r.getparent()
        context_text = own_text
        if parent is not None and parent.tag in MATH_PARENTS:
            context_text = ''
            for child in parent:
                if child.tag == f'{M}r':
                    for mt in child.iter(f'{M}t'):
                        if mt.text:
                            context_text += mt.text
        context_text = context_text.strip()

        # ── Classify based on own_text first, context only for disambiguation ──
        # ISO 80000-2: numbers ALWAYS upright; single-letter variables italic;
        # multi-letter identifiers (functions, descriptive) upright.
        need_upright = False

        if own_text in UPRIGHT_CONSTANTS:
            need_upright = True
        elif re.fullmatch(r'[\d.+\-′″]+', own_text):
            need_upright = True          # numbers always upright
        elif re.fullmatch(r'[A-Za-z]+[0-9]+', own_text) or re.fullmatch(r'[0-9]+[A-Za-z]+', own_text):
            need_upright = True          # alphanumeric: HE1, DPf100, 3a
        elif re.fullmatch(r'[A-Za-zΑ-ω]{2,}', own_text):
            need_upright = True          # multi-letter: function / descriptive
        elif len(own_text) == 1 and (own_text.isalpha() or
            ('Α' <= own_text <= 'ω')):
            # Single letter — check context: if adjacent siblings make multi-letter, go upright
            if len(context_text) >= 2 and re.fullmatch(r'[A-Za-zΑ-ω]{2,}', context_text):
                need_upright = True      # actually part of a multi-letter word
            # else: single-letter variable → italic (default)
        else:
            pass                         # mixed / punctuation → keep default

        sty = m_rPr.find(f'{M}sty')
        if need_upright:
            if sty is None:
                sty = etree.SubElement(m_rPr, f'{M}sty')
            sty.set(f'{M}val', 'p')
            upright += 1
        else:
            if sty is not None:
                m_rPr.remove(sty)
            italic += 1

    print(f"  Math font: {total} runs → TNR ({italic} italic, {upright} upright)")

    # Also fix settings.xml
    settings_path = os.path.join(tmp_dir, 'word', 'settings.xml')
    if os.path.exists(settings_path):
        settings_tree = etree.parse(settings_path)
        settings_root = settings_tree.getroot()
        math_pr = settings_root.find(f'{W}mathPr')
        if math_pr is None:
            math_pr = etree.SubElement(settings_root, f'{W}mathPr')
        math_font = math_pr.find(f'{W}mathFont')
        if math_font is None:
            math_font = etree.SubElement(math_pr, f'{W}mathFont')
        math_font.set(f'{W}val', 'Times New Roman')
        settings_new = etree.tostring(settings_root, encoding='UTF-8', xml_declaration=True, standalone=None)
        with open(settings_path, 'wb') as f:
            f.write(settings_new)
        print("  Math font: default set in settings.xml")


# ── Header / Footer (第十二条) ─────────────────────────────────
def add_header_footer(tmp_dir, doc_root):
    """Add header (隶书三号居中) and footer (page numbers) to the docx.

    Header: "武汉科技大学本科毕业论文", 隶书三号(16pt), centered
    Footer: Page number, Times New Roman 小五号(9pt), centered
    """
    rels_dir = os.path.join(tmp_dir, 'word', '_rels')
    os.makedirs(rels_dir, exist_ok=True)

    # Check existing rels count
    header_dir = os.path.join(tmp_dir, 'word')
    existing_headers = [f for f in os.listdir(header_dir) if f.startswith('header') and f.endswith('.xml')]
    existing_footers = [f for f in os.listdir(header_dir) if f.startswith('footer') and f.endswith('.xml')]

    hdr_num = len(existing_headers) + 1
    ftr_num = len(existing_footers) + 1
    hdr_file = f'header{hdr_num}.xml'
    ftr_file = f'footer{ftr_num}.xml'

    # ── Create header XML ──
    # Build with lxml for proper namespaces
    hdr_nsmap = {
        'w': W_NS,
        'r': R_NS,
        'mc': MC_NS,
        'wps': WP_NS,
        'v': V_NS,
        'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': PIC_NS,
        'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
        'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    }

    # Simple header with single paragraph
    hdr_root = etree.Element(f'{W}hdr', nsmap={'w': W_NS, 'r': R_NS})

    hdr_p = etree.SubElement(hdr_root, f'{W}p')
    hdr_pPr = etree.SubElement(hdr_p, f'{W}pPr')
    hdr_jc = etree.SubElement(hdr_pPr, f'{W}jc')
    hdr_jc.set(f'{W}val', 'center')
    hdr_ind = etree.SubElement(hdr_pPr, f'{W}ind')
    hdr_ind.set(f'{W}firstLine', '0')
    hdr_ind.set(f'{W}firstLineChars', '0')
    hdr_spacing = etree.SubElement(hdr_pPr, f'{W}spacing')
    hdr_spacing.set(f'{W}line', '240')
    hdr_spacing.set(f'{W}lineRule', 'auto')

    hdr_r = etree.SubElement(hdr_p, f'{W}r')
    hdr_rPr = etree.SubElement(hdr_r, f'{W}rPr')
    hdr_fonts = etree.SubElement(hdr_rPr, f'{W}rFonts')
    hdr_fonts.set(f'{W}ascii', '隶书')
    hdr_fonts.set(f'{W}hAnsi', '隶书')
    hdr_fonts.set(f'{W}eastAsia', '隶书')
    hdr_sz = etree.SubElement(hdr_rPr, f'{W}sz')
    hdr_sz.set(f'{W}val', '32')  # 三号 = 16pt → sz=32
    hdr_szCs = etree.SubElement(hdr_rPr, f'{W}szCs')
    hdr_szCs.set(f'{W}val', '32')
    hdr_t = etree.SubElement(hdr_r, f'{W}t')
    hdr_t.set(f'{{{W_NS}}}space', 'preserve')
    hdr_t.text = HEADER_TEXT

    hdr_xml = etree.tostring(hdr_root, encoding='UTF-8', xml_declaration=True, standalone=None)
    with open(os.path.join(tmp_dir, 'word', hdr_file), 'wb') as f:
        f.write(hdr_xml)
    print(f"  Header created: {hdr_file} (隶书三号, centered)")

    # ── Create footer XML ──
    ftr_root = etree.Element(f'{W}ftr', nsmap={'w': W_NS, 'r': R_NS})

    ftr_p = etree.SubElement(ftr_root, f'{W}p')
    ftr_pPr = etree.SubElement(ftr_p, f'{W}pPr')
    ftr_ind = etree.SubElement(ftr_pPr, f'{W}ind')
    ftr_ind.set(f'{W}firstLine', '0')
    ftr_ind.set(f'{W}firstLineChars', '0')
    ftr_jc = etree.SubElement(ftr_pPr, f'{W}jc')
    ftr_jc.set(f'{W}val', 'center')

    ftr_r = etree.SubElement(ftr_p, f'{W}r')
    ftr_rPr = etree.SubElement(ftr_r, f'{W}rPr')
    ftr_fonts = etree.SubElement(ftr_rPr, f'{W}rFonts')
    ftr_fonts.set(f'{W}ascii', 'Times New Roman')
    ftr_fonts.set(f'{W}hAnsi', 'Times New Roman')
    ftr_fonts.set(f'{W}eastAsia', '宋体')
    ftr_sz = etree.SubElement(ftr_rPr, f'{W}sz')
    ftr_sz.set(f'{W}val', '20')  # 小五号 = 9pt → sz=20
    ftr_szCs = etree.SubElement(ftr_rPr, f'{W}szCs')
    ftr_szCs.set(f'{W}val', '20')

    # Page number field: <w:fldChar w:fldCharType="begin"/> <w:instrText> PAGE </w:instrText> ...
    fld_begin = etree.SubElement(ftr_p, f'{W}r')
    fld_begin_rPr = etree.SubElement(fld_begin, f'{W}rPr')
    fld_begin_fonts = etree.SubElement(fld_begin_rPr, f'{W}rFonts')
    fld_begin_fonts.set(f'{W}ascii', 'Times New Roman')
    fld_begin_fonts.set(f'{W}hAnsi', 'Times New Roman')
    fld_begin_sz = etree.SubElement(fld_begin_rPr, f'{W}sz')
    fld_begin_sz.set(f'{W}val', '20')
    fld_begin_ch = etree.SubElement(fld_begin, f'{W}fldChar')
    fld_begin_ch.set(f'{W}fldCharType', 'begin')

    instr = etree.SubElement(ftr_p, f'{W}r')
    instr_rPr = etree.SubElement(instr, f'{W}rPr')
    instr_fonts = etree.SubElement(instr_rPr, f'{W}rFonts')
    instr_fonts.set(f'{W}ascii', 'Times New Roman')
    instr_fonts.set(f'{W}hAnsi', 'Times New Roman')
    instr_sz = etree.SubElement(instr_rPr, f'{W}sz')
    instr_sz.set(f'{W}val', '20')
    instr_text = etree.SubElement(instr, f'{W}instrText')
    instr_text.set(f'{{{W_NS}}}space', 'preserve')
    instr_text.text = ' PAGE '

    fld_end = etree.SubElement(ftr_p, f'{W}r')
    fld_end_rPr = etree.SubElement(fld_end, f'{W}rPr')
    fld_end_fonts = etree.SubElement(fld_end_rPr, f'{W}rFonts')
    fld_end_fonts.set(f'{W}ascii', 'Times New Roman')
    fld_end_fonts.set(f'{W}hAnsi', 'Times New Roman')
    fld_end_sz = etree.SubElement(fld_end_rPr, f'{W}sz')
    fld_end_sz.set(f'{W}val', '20')
    fld_end_ch = etree.SubElement(fld_end, f'{W}fldChar')
    fld_end_ch.set(f'{W}fldCharType', 'end')

    ftr_xml = etree.tostring(ftr_root, encoding='UTF-8', xml_declaration=True, standalone=None)
    with open(os.path.join(tmp_dir, 'word', ftr_file), 'wb') as f:
        f.write(ftr_xml)
    print(f"  Footer created: {ftr_file} (page numbers, 小五号 TNR, centered)")

    # ── Add relationships in document.xml.rels ──
    doc_rels_path = os.path.join(tmp_dir, 'word', '_rels', 'document.xml.rels')
    rels_tree = etree.parse(doc_rels_path)
    rels_root = rels_tree.getroot()

    rels_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    rel_ns = f'{{{rels_ns}}}'

    # Find existing max rId
    existing_ids = []
    for rel in rels_root.iter(f'{rel_ns}Relationship'):
        rid = rel.get('Id')
        if rid:
            existing_ids.append(int(rid.replace('rId', '')))
    next_id = max(existing_ids) + 1 if existing_ids else 1

    # Add header relationship
    hdr_rel = etree.SubElement(rels_root, f'{rel_ns}Relationship')
    hdr_rel.set('Id', f'rId{next_id}')
    hdr_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header')
    hdr_rel.set('Target', hdr_file)
    hdr_rid = f'rId{next_id}'
    next_id += 1

    # Add footer relationship
    ftr_rel = etree.SubElement(rels_root, f'{rel_ns}Relationship')
    ftr_rel.set('Id', f'rId{next_id}')
    ftr_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
    ftr_rel.set('Target', ftr_file)
    ftr_rid = f'rId{next_id}'

    rels_tree.write(doc_rels_path, encoding='UTF-8', xml_declaration=True)

    # ── Add header/footer references to sectPr in document.xml ──
    for sectPr in doc_root.iter(f'{W}sectPr'):
        # Remove any existing header/footer refs
        for old_hdr in sectPr.findall(f'{W}headerReference'):
            sectPr.remove(old_hdr)
        for old_ftr in sectPr.findall(f'{W}footerReference'):
            sectPr.remove(old_ftr)

        # Add header reference
        hdr_ref = etree.Element(f'{W}headerReference')
        hdr_ref.set(f'{W}type', 'default')
        hdr_ref.set(f'{R}id', hdr_rid)
        # Insert as first child of sectPr (before pgSz)
        sectPr.insert(0, hdr_ref)

        # Add footer reference
        ftr_ref = etree.Element(f'{W}footerReference')
        ftr_ref.set(f'{W}type', 'default')
        ftr_ref.set(f'{R}id', ftr_rid)
        sectPr.insert(1, ftr_ref)

        # Add titlePg if not present (suppress header/footer on first page - for TOC)
        title_pg = etree.SubElement(sectPr, f'{W}titlePg')

        break  # Only first section

    print(f"  Header/footer refs added to document.xml (rIds: {hdr_rid}, {ftr_rid})")

    # ── Update [Content_Types].xml ──
    content_types_path = os.path.join(tmp_dir, '[Content_Types].xml')
    ct_tree = etree.parse(content_types_path)
    ct_root = ct_tree.getroot()
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    ct_n = f'{{{ct_ns}}}'

    # Add Override for header
    hdr_override = etree.SubElement(ct_root, f'{ct_n}Override')
    hdr_override.set('PartName', f'/word/{hdr_file}')
    hdr_override.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')

    # Add Override for footer
    ftr_override = etree.SubElement(ct_root, f'{ct_n}Override')
    ftr_override.set('PartName', f'/word/{ftr_file}')
    ftr_override.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')

    ct_tree.write(content_types_path, encoding='UTF-8', xml_declaration=True)
    print("  [Content_Types].xml updated")


# ── Main ──────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("WUST Thesis Formatting Script")
    print("Based on: 武汉科技大学本科毕业设计(论文)文件汇编(2024版)")
    print("=" * 60)

    # Step 1: Fix reference styles
    fixed_ref = fix_reference_styles()

    # Step 2: Pandoc conversion
    raw_docx = run_pandoc(fixed_ref)

    # Step 3: Post-process
    post_process(raw_docx)

    print("=" * 60)
    print(f"Done: {OUTPUT}")
    print("=" * 60)


if __name__ == '__main__':
    main()
