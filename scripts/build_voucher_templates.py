#!/usr/bin/env python3
"""Build docxtemplater templates from real voucher samples, keeping styles/media."""
from __future__ import annotations
import copy, re, shutil, zipfile
from pathlib import Path
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def q(tag): return f'{W}{tag}'

def para_text(p):
  return ''.join(t.text or '' for t in p.findall(f'.//{q("t")}'))

def clean_rPr(rPr):
  if rPr is None: return None
  rPr = copy.deepcopy(rPr)
  for tag in ('u', 'color', 'highlight'):
    for el in list(rPr.findall(q(tag))):
      rPr.remove(el)
  return rPr

def set_para_placeholder(p, placeholder, keep_tabs_before=None):
  """Replace paragraph content with a single run containing placeholder, preserving first rPr/pPr."""
  pPr = p.find(q('pPr'))
  if pPr is None:
    pPr = etree.Element(q('pPr')); p.insert(0, pPr)
  for tabs in list(pPr.findall(q('tabs'))):
    pPr.remove(tabs)
  jc = pPr.find(q('jc'))
  if jc is None: jc = etree.SubElement(pPr, q('jc'))
  jc.set(f'{W}val', 'left')
  pr = pPr.find(q('rPr'))
  if pr is not None:
    for tag in ('u', 'highlight'):
      for el in list(pr.findall(q(tag))): pr.remove(el)
  first_rPr = None
  for r in p.findall(q('r')):
    rp = r.find(q('rPr'))
    if rp is not None:
      first_rPr = clean_rPr(rp)
      break
  for child in list(p):
    if child is not pPr:
      p.remove(child)
  # optional tab runs before placeholder (for UpDate/Payment alignment)
  if keep_tabs_before:
    for _ in range(keep_tabs_before):
      r = etree.SubElement(p, q('r'))
      if first_rPr is not None:
        r.append(copy.deepcopy(first_rPr))
      etree.SubElement(r, q('tab'))
  r = etree.SubElement(p, q('r'))
  if first_rPr is not None:
    r.append(copy.deepcopy(first_rPr))
  t = etree.SubElement(r, q('t'))
  t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
  t.text = placeholder

def replace_in_para_keep_label(p, label_prefix, placeholder):
  """Keep label text (e.g. 'UpDate:') and tabs, replace value with placeholder."""
  texts = [(t, t.text or '') for t in p.findall(f'.//{q("t")}')]
  full = ''.join(x[1] for x in texts)
  if label_prefix not in full and not full.strip().startswith(label_prefix.replace(':','')):
    return False
  pPr = p.find(q('pPr'))
  # Collect leading label + tabs from runs
  first_rPr = None
  for r in p.findall(q('r')):
    rp = r.find(q('rPr'))
    if rp is not None:
      first_rPr = copy.deepcopy(rp)
      break
  # Rebuild: label run + tabs + placeholder
  for child in list(p):
    if child is not pPr:
      p.remove(child)
  # label
  r = etree.SubElement(p, q('r'))
  if first_rPr is not None:
    r.append(copy.deepcopy(first_rPr))
  t = etree.SubElement(r, q('t'))
  t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
  t.text = label_prefix
  # sample uses tabs between label and value — add 3 tabs like original feel
  for _ in range(3):
    r = etree.SubElement(p, q('r'))
    if first_rPr is not None:
      r.append(copy.deepcopy(first_rPr))
    etree.SubElement(r, q('tab'))
  r = etree.SubElement(p, q('r'))
  if first_rPr is not None:
    r.append(copy.deepcopy(first_rPr))
  t = etree.SubElement(r, q('t'))
  t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
  t.text = placeholder
  return True


def set_cell_text(tc, text, bold=None):
  """Replace all paragraphs in cell with one paragraph / one run."""
  # keep tcPr
  tcPr = tc.find(q('tcPr'))
  # get pPr and rPr from first paragraph/run
  first_p = tc.find(q('p'))
  pPr = copy.deepcopy(first_p.find(q('pPr'))) if first_p is not None and first_p.find(q('pPr')) is not None else None
  if pPr is not None:
    pr = pPr.find(q('rPr'))
    if pr is not None:
      cleaned = clean_rPr(pr)
      pPr.remove(pr)
      if cleaned is not None: pPr.append(cleaned)
  first_rPr = None
  if first_p is not None:
    for r in first_p.findall(q('r')):
      if r.find(q('rPr')) is not None:
        first_rPr = clean_rPr(r.find(q('rPr')))
        break
  for child in list(tc):
    if child is not tcPr:
      tc.remove(child)
  # support multi-line text -> multiple paragraphs (except placeholders with \n use linebreak in one p for docxtemplater)
  p = etree.SubElement(tc, q('p'))
  if pPr is not None:
    p.insert(0, pPr)
  r = etree.SubElement(p, q('r'))
  if first_rPr is not None:
    rPr = copy.deepcopy(first_rPr)
    if bold is True:
      for tag in ('b','bCs'):
        el = rPr.find(q(tag))
        if el is None:
          el = etree.SubElement(rPr, q(tag))
        el.set(q('val') if False else f'{W}val', '1')
    r.append(rPr)
  elif bold:
    rPr = etree.SubElement(r, q('rPr'))
    etree.SubElement(rPr, q('b')).set(f'{W}val', '1')
  t = etree.SubElement(r, q('t'))
  t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
  t.text = text

def prepend_to_cell(tc, prefix):
  """Prepend text to first w:t in cell (for loop tags)."""
  ts = tc.findall(f'.//{q("t")}')
  if not ts:
    set_cell_text(tc, prefix)
    return
  t0 = ts[0]
  t0.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
  t0.text = prefix + (t0.text or '')

def append_to_cell(tc, suffix):
  ts = tc.findall(f'.//{q("t")}')
  if not ts:
    set_cell_text(tc, suffix)
    return
  tN = ts[-1]
  tN.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
  tN.text = (tN.text or '') + suffix

def build_accommodation(src: Path, dst: Path):
  tmp = Path('/tmp/tpl_acc_build')
  if tmp.exists(): shutil.rmtree(tmp)
  tmp.mkdir()
  with zipfile.ZipFile(src) as z: z.extractall(tmp)
  doc_path = tmp/'word'/'document.xml'
  tree = etree.parse(str(doc_path))
  body = tree.getroot().find(f'.//{q("body")}')

  # --- header paragraphs ---
  paras = [c for c in body if c.tag == q('p')]
  # Identify by content
  guest_paras = []
  for p in paras:
    txt = para_text(p)
    if txt.strip().startswith('VOUCHER'):
      pass  # keep title
    elif 'UpDate' in txt or 'UpDate:' in txt.replace(' ',''):
      replace_in_para_keep_label(p, 'UpDate:', '{update_date}')
    elif txt.startswith('Payment') or 'Payment:' in txt:
      replace_in_para_keep_label(p, 'Payment:', '{payment}')
    elif re.match(r'\s*\d+\.\s+', txt) or re.match(r'\d+\.\s+MR', txt.strip()):
      guest_paras.append(p)
    elif txt.startswith('Meal') or 'Meal' in txt[:6]:
      replace_in_para_keep_label(p, 'Meal', '{meal}')
      # Meal label might need tab - fix to "Meal" + tabs + value. Sample is "Meal\t\tBreakfast"
      # redo more carefully
      set_para_placeholder(p, 'Meal')
      # actually rebuild Meal\t\t{meal}
      pPr = p.find(q('pPr'))
      for child in list(p):
        if child is not pPr: p.remove(child)
      r=etree.SubElement(p,q('r')); t=etree.SubElement(r,q('t')); t.text='Meal'
      for _ in range(4):
        r=etree.SubElement(p,q('r')); etree.SubElement(r,q('tab'))
      r=etree.SubElement(p,q('r')); t=etree.SubElement(r,q('t')); t.set('{http://www.w3.org/XML/1998/namespace}space','preserve'); t.text='{meal}'
    elif 'Number of passengers' in txt:
      replace_in_para_keep_label(p, 'Number of passengers:', '{pax}')
    elif 'Contact for delegate' in txt:
      replace_in_para_keep_label(p, 'Contact for delegate: ', '{contact}')

  # Collapse guest paragraphs into one {guest_lines}
  if guest_paras:
    first = guest_paras[0]
    set_para_placeholder(first, '{guest_lines}')
    pPr = first.find(q('pPr'))
    if pPr is None:
      pPr = etree.Element(q('pPr')); first.insert(0, pPr)
    jc = pPr.find(q('jc'))
    if jc is None: jc = etree.SubElement(pPr, q('jc'))
    jc.set(f'{W}val', 'left')
    for p in guest_paras[1:]:
      parent = p.getparent()
      parent.remove(p)

  # --- table ---
  tbl = body.find(q('tbl'))
  rows = tbl.findall(q('tr'))
  # Keep: header row r0, data row r1, train row r6 as prototypes
  header_row = copy.deepcopy(rows[0])
  data_row = copy.deepcopy(rows[1])
  train_row = copy.deepcopy(rows[6])

  # Clear private data from prototypes -> placeholders
  # header row: keep CITY etc labels, add loop/condition tags
  hcells = header_row.findall(q('tc'))
  # Ensure header labels clean
  labels = ['CITY','CHECK-IN','CHECK-OUT','HOTEL','NOTE / BOOKING CODE']
  for i,lab in enumerate(labels):
    set_cell_text(hcells[i], lab)

  dcells = data_row.findall(q('tc'))
  set_cell_text(dcells[0], '{city}')
  set_cell_text(dcells[1], '{check_in}')
  set_cell_text(dcells[2], '{check_out}')
  set_cell_text(dcells[3], '{hotel_block}')
  set_cell_text(dcells[4], '{note}')

  tcells = train_row.findall(q('tc'))
  set_cell_text(tcells[0], '{train_text}')

  # Build repeating block:
  # {#acc_rows}
  #   {#is_hotel}{#show_header} header {/show_header} data {/is_hotel}
  #   {#is_train} train {/is_train}
  # {/acc_rows}

  # Clone rows and add tags
  hdr = copy.deepcopy(header_row)
  dat = copy.deepcopy(data_row)
  trn = copy.deepcopy(train_row)

  # Nested conditionals for hotel with optional header:
  # Row hdr: {#acc_rows}{#is_hotel}{#show_header} CITY ... {/show_header}
  # But {/show_header} must close on same row. Then data row ends with {/is_hotel}
  # Train: {#is_train}...{/is_train}{/acc_rows}

  # Better structure used by docxtemplater for optional rows:
  # header row: {#acc_rows}{#is_hotel}{#show_header} ... {/show_header}{/is_hotel}{^is_hotel}{#is_train}PLACEHOLDER{/is_train}{/} 
  # That's getting messy.

  # Simpler approach that works: flatten in JS to sections. OR use:
  # 
  # {#acc_rows}{#is_hotel}
  #   header row (always)  
  #   data row
  # {/is_hotel}{/acc_rows}
  # AND separately we accept header before Hue (minor vs sample) 
  #
  # User asked for exact look including merged train. Extra header after train is a small deviation.
  # OR: show_header conditional row:
  #
  # {#acc_rows}
  # {#is_hotel}{#show_header}
  # [header cells...]
  # {/show_header}
  # [data cells...]
  # {/is_hotel}
  # {#is_train}
  # [train]
  # {/is_train}
  # {/acc_rows}
  #
  # For this, show_header open must be in header row first cell, close in header row last cell.
  # is_hotel open before show_header in header row — but if show_header is false, header row is removed,
  # and is_hotel open tag would be removed too! That's a problem.
  #
  # Docxtemplater doc: put {#is_hotel} on data row if header is optional via separate flag rows.
  #
  # Use raw XML approach: always output 2 rows for hotel (header+data), 1 for train.
  # For post-train hotel, still show header (green) — closer to consistency; sample omitted it
  # maybe by accident. Checking sample Hue - no header. I'll use show_header with a trick:
  #
  # Use ONE loop over prebuilt `acc_blocks` where each block is already XML-like data —
  # actually simplest fix: in JS expand to `acc_print_rows`: list of {row_type, ...}` 
  # and template has three conditional single-row types.
  #
  # {#acc_print_rows}
  # {#is_hdr} header labels {/is_hdr}
  # {#is_data} city checkin... {/is_data}
  # {#is_train} train_text {/is_train}
  # {/acc_print_rows}
  #
  # Each iteration is exactly one table row. Perfect for docxtemplater.

  # We'll change JS to build acc_print_rows. Template uses that.

  hdr2 = copy.deepcopy(header_row)
  dat2 = copy.deepcopy(data_row)
  trn2 = copy.deepcopy(train_row)
  hcells = hdr2.findall(q('tc'))
  dcells = dat2.findall(q('tc'))
  tcells = trn2.findall(q('tc'))

  # Three prototype rows inside one loop — use conditionals so only one shows per iteration
  # Row1 (header type): {#acc_print_rows}{#is_hdr} CITY ... {/is_hdr}{#is_data}{city}... 
  # Actually each row element is separate — can't put three row types in one iteration easily
  # without all three rows existing in the template loop body.

  # CORRECT docxtemplater pattern for multi-type rows:
  # Loop body contains ALL row variants; conditionals empty the ones that don't apply.
  # 
  # {#acc_print_rows}
  #   {#is_hdr}<w:tr>header</w:tr>{/is_hdr}
  #   {#is_data}<w:tr>data</w:tr>{/is_data}
  #   {#is_train}<w:tr>train</w:tr>{/is_train}
  # {/acc_print_rows}
  #
  # When using table rows, put tags inside first/last cells:
  # https://docxtemplater.com/docs/faq/#loop-table
  #
  # For conditional rows, use:
  # {#is_hdr} in first cell, {/is_hdr} in last cell of that row — row removed if false.

  prepend_to_cell(hcells[0], '{#acc_print_rows}{#is_hdr}')
  append_to_cell(hcells[-1], '{/is_hdr}')

  prepend_to_cell(dcells[0], '{#is_data}')
  append_to_cell(dcells[-1], '{/is_data}')

  prepend_to_cell(tcells[0], '{#is_train}')
  append_to_cell(tcells[-1], '{/is_train}{/acc_print_rows}')

  # Remove all existing rows and add the three prototype rows
  for r in tbl.findall(q('tr')):
    tbl.remove(r)
  tbl.append(hdr2)
  tbl.append(dat2)
  tbl.append(trn2)

  tree.write(str(doc_path), xml_declaration=True, encoding='UTF-8', standalone=True)

  # Repack
  if dst.exists(): dst.unlink()
  with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
    for f in tmp.rglob('*'):
      if f.is_file():
        z.write(f, f.relative_to(tmp).as_posix())
  print('Wrote', dst, 'size', dst.stat().st_size)

def build_transfers(src: Path, dst: Path):
  tmp = Path('/tmp/tpl_trf_build')
  if tmp.exists(): shutil.rmtree(tmp)
  tmp.mkdir()
  with zipfile.ZipFile(src) as z: z.extractall(tmp)
  doc_path = tmp/'word'/'document.xml'
  tree = etree.parse(str(doc_path))
  body = tree.getroot().find(f'.//{q("body")}')

  paras = [c for c in body if c.tag == q('p')]
  guest_paras = []
  for p in paras:
    txt = para_text(p)
    if 'UpDate' in txt:
      replace_in_para_keep_label(p, 'UpDate:', '{update_date}')
    elif txt.startswith('Payment') or 'Payment:' in txt:
      replace_in_para_keep_label(p, 'Payment:', '{payment}')
    elif re.match(r'\s*\d+\.\s+', txt) or re.match(r'\d+\.\s+MR', txt.strip()):
      guest_paras.append(p)
    elif 'Number of passengers' in txt:
      replace_in_para_keep_label(p, 'Number of passengers:', '{pax}')
    elif 'Contact for delegate' in txt:
      replace_in_para_keep_label(p, 'Contact for delegate: ', '{contact}')
  if guest_paras:
    first = guest_paras[0]
    set_para_placeholder(first, '{guest_lines}')
    pPr = first.find(q('pPr'))
    if pPr is None:
      pPr = etree.Element(q('pPr')); first.insert(0, pPr)
    jc = pPr.find(q('jc'))
    if jc is None: jc = etree.SubElement(pPr, q('jc'))
    jc.set(f'{W}val', 'left')
    for p in guest_paras[1:]:
      p.getparent().remove(p)

  tbl = body.find(q('tbl'))
  rows = tbl.findall(q('tr'))
  header_row = copy.deepcopy(rows[0])
  data_row = copy.deepcopy(rows[1])

  hcells = header_row.findall(q('tc'))
  for i,lab in enumerate(['TRANSFER','DATE','WHERE FROM','WHERE TO','NOTE']):
    set_cell_text(hcells[i], lab)

  dcells = data_row.findall(q('tc'))
  set_cell_text(dcells[0], '{label}')
  set_cell_text(dcells[1], '{date}')
  set_cell_text(dcells[2], '{where_from}')
  set_cell_text(dcells[3], '{where_to}')
  set_cell_text(dcells[4], '{note}')

  # Transfers sample: each item is header+data pair. Use trf_print_rows with is_hdr/is_data
  hdr2 = copy.deepcopy(header_row)
  dat2 = copy.deepcopy(data_row)
  hcells = hdr2.findall(q('tc'))
  dcells = dat2.findall(q('tc'))
  prepend_to_cell(hcells[0], '{#trf_print_rows}{#is_hdr}')
  append_to_cell(hcells[-1], '{/is_hdr}')
  prepend_to_cell(dcells[0], '{#is_data}')
  append_to_cell(dcells[-1], '{/is_data}{/trf_print_rows}')

  for r in tbl.findall(q('tr')):
    tbl.remove(r)
  tbl.append(hdr2)
  tbl.append(dat2)

  tree.write(str(doc_path), xml_declaration=True, encoding='UTF-8', standalone=True)
  if dst.exists(): dst.unlink()
  with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
    for f in tmp.rglob('*'):
      if f.is_file():
        z.write(f, f.relative_to(tmp).as_posix())
  print('Wrote', dst, 'size', dst.stat().st_size)

# Verify no private data remains
PRIVATE = ['LEKES','David Vu','94 455','ORIENTAL','ASPIRA','TAM COC','CHERISH','HOI AN BEACH','ODYS','VN036','VN 036','TH.140626','127095','469509','55747','28.04.2026']

def check_clean(path):
  with zipfile.ZipFile(path) as z:
    xml = z.read('word/document.xml').decode('utf-8', errors='ignore')
  bad = [p for p in PRIVATE if p.lower() in xml.lower()]
  # Allow hotel names? No - strip all guest/private. Hotel names in template placeholders only.
  print(path.name, 'private hits:', bad)
  return bad

if __name__ == '__main__':
  src_a = Path('/workspace/docgen-real/voucher_a.docx')
  src_b = Path('/workspace/docgen-real/voucher_b.docx')
  dst_a = Path('/workspace/docgen-web/templates/voucher_accommodation.docx')
  dst_b = Path('/workspace/docgen-web/templates/voucher_transfers.docx')
  build_accommodation(src_a, dst_a)
  build_transfers(src_b, dst_b)
  check_clean(dst_a)
  check_clean(dst_b)
