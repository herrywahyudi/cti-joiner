import streamlit as st
import pandas as pd
import re, io, zipfile
from datetime import datetime
from openpyxl import load_workbook

st.set_page_config(page_title="CTI Joiner Report", page_icon="📊", layout="wide")

# ── Helpers ───────────────────────────────────────────────────────────────────
def norm(text):
    return str(text).strip().lower().replace('_',' ').replace('-',' ').replace('/',' ')

def norm_id(v):
    return str(v).lower().replace('-','').replace('_','').replace(' ','')

def find_col(df, names):
    nc = {norm(c): c for c in df.columns}
    for n in names:
        if norm(n) in nc: return nc[norm(n)]
    return None

def clean_cell(v):
    v = str(v).strip()
    return '' if v.lower() == 'nan' else v

def clean_id(v):
    v = clean_cell(str(v))
    if not v: return ''
    try:
        n = float(v)
        if n.is_integer(): return str(int(n))
    except: pass
    return v[:-2] if v.endswith('.0') else v

JOINER_E_KEYS = [
    'enumbercode', 'enumber',
    'seafareridnumber', 'seafarerid',
    'employeeid', 'crewid',
]

CTI_OFFICES = [
    'CTI Indonesia',
    'CTI Group Bangkok',
    'CTI Group MCSI',
    'CTI Group Myanmar',
    'CTI Group South Africa',
    'CTI Group Vietnam',
    'CTI Partner Kendrick',
]

STATUSES = ['New Hire', 'Repeater', 'Re Hire', 'Resigned']

# ── Core functions ────────────────────────────────────────────────────────────
def build_lookup_dict(zoho_bytes, inactive_bytes):
    lookup = {}
    for label, data in [('inactive', inactive_bytes), ('zoho', zoho_bytes)]:
        if not data: continue
        try:
            df = pd.read_excel(io.BytesIO(data))
            e_col   = find_col(df, ['Seafarer ID Number','Seafarer ID','E-Number Code','Crew ID'])
            cti_col = find_col(df, ['CTI Office'])
            sta_col = find_col(df, ['Employment Status'])
            if not e_col: continue
            for _, row in df.dropna(subset=[e_col]).iterrows():
                e = clean_id(row[e_col])
                if e:
                    lookup[e] = {
                        'cti':    clean_cell(row[cti_col]) if cti_col else '',
                        'status': clean_cell(row[sta_col]) if sta_col else '',
                    }
        except: pass
    return lookup

def run_joiner_fill(joiner_bytes, lookup_dict, manual_overrides):
    wb = load_workbook(io.BytesIO(joiner_bytes))
    log_lines = []
    total_filled = 0
    total_missing = []

    for sh in wb.sheetnames:
        ws = wb[sh]
        if ws.max_row < 2: continue

        headers = {}
        for c in range(1, ws.max_column + 1):
            val = ws.cell(1, c).value
            if val: headers[norm_id(str(val))] = c

        e_col = None
        for k in JOINER_E_KEYS:
            if k in headers:
                e_col = headers[k]
                break

        if not e_col:
            log_lines.append(f"!! Sheet '{sh}': no E-Number column, skipped")
            continue

        cti_col_idx  = headers.get(norm_id('CTI Office'))
        stat_col_idx = headers.get(norm_id('Employment Status'))

        if not cti_col_idx:
            ins = e_col + 1
            ws.insert_cols(ins)
            ws.cell(1, ins).value = 'CTI Office'
            cti_col_idx = ins
            if stat_col_idx and stat_col_idx >= ins:
                stat_col_idx += 1

        if not stat_col_idx:
            ins = cti_col_idx + 1
            ws.insert_cols(ins)
            ws.cell(1, ins).value = 'Employment Status'
            stat_col_idx = ins

        filled  = 0
        missing = []

        for r in range(2, ws.max_row + 1):
            e = clean_id(ws.cell(r, e_col).value)
            if not e: continue

            if e in manual_overrides:
                ov = manual_overrides[e]
                if ov.get('cti'):    ws.cell(r, cti_col_idx).value  = ov['cti']
                if ov.get('status'): ws.cell(r, stat_col_idx).value = ov['status']
                filled += 1
                continue

            if e in lookup_dict:
                d = lookup_dict[e]
                if d['cti']:    ws.cell(r, cti_col_idx).value  = d['cti']
                if d['status']: ws.cell(r, stat_col_idx).value = d['status']
                filled += 1
            else:
                name_col = headers.get(norm_id('name')) or headers.get(norm_id('full name'))
                name = clean_cell(ws.cell(r, name_col).value) if name_col else ''
                missing.append({'E-Number': e, 'Name': name, 'Sheet': sh})

        log_lines.append(f"OK  Sheet '{sh}': {filled} filled, {len(missing)} not found in master")
        total_filled += filled
        total_missing.extend(missing)

    log_lines.append(f"\nTotal: {total_filled} filled | {len(total_missing)} missing")
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), '\n'.join(log_lines), total_missing

def run_split(filled_bytes):
    xls  = pd.ExcelFile(io.BytesIO(filled_bytes))
    books = {}
    tag  = datetime.now().strftime('%Y%m%d')

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        if 'CTI Office' not in df.columns: continue
        for office, sub in df.groupby(df['CTI Office'].fillna('UNKNOWN')):
            office = str(office).strip().replace('/', '-')
            if office.upper() == 'UNKNOWN': continue
            books.setdefault(office, {})[sheet] = sub

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for office, sheets in books.items():
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                for sname, data in sheets.items():
                    data.to_excel(writer, sheet_name=sname, index=False)
            buf.seek(0)
            zf.writestr(f'{office}_{tag}.xlsx', buf.read())

    zip_buf.seek(0)
    return zip_buf.getvalue(), sorted(books.keys())

# ── UI ────────────────────────────────────────────────────────────────────────
st.title('📊 CTI Joiner Report Tool')
st.caption('Fill CTI Office & Employment Status · Split by office · CTI Indonesia — Bali')

# Step 1
st.header('Step 1 — Upload Files')
c1, c2, c3 = st.columns(3)
with c1:
    joiner_file   = st.file_uploader('Joiner Report', type=['xlsx','xls'], key='joiner')
with c2:
    zoho_file     = st.file_uploader('Zoho Master', type=['xlsx','xls'], key='zoho')
with c3:
    inactive_file = st.file_uploader('Inactive Master (optional)', type=['xlsx','xls'], key='inactive')

st.divider()

# Step 2
st.header('Step 2 — Manual Overrides')
st.caption('For E-Numbers not found in the master — fill them in here before running.')

if 'overrides' not in st.session_state:
    st.session_state['overrides'] = [{'e': '', 'cti': CTI_OFFICES[0], 'status': STATUSES[0]}]

for i, ov in enumerate(st.session_state['overrides']):
    c1, c2, c3, c4 = st.columns([2, 2, 2, 0.5])
    with c1:
        st.session_state['overrides'][i]['e'] = st.text_input(
            'E-Number', value=ov['e'], key=f'ov_e_{i}', placeholder='e.g. 857218')
    with c2:
        idx = CTI_OFFICES.index(ov['cti']) if ov['cti'] in CTI_OFFICES else 0
        st.session_state['overrides'][i]['cti'] = st.selectbox(
            'CTI Office', CTI_OFFICES, index=idx, key=f'ov_cti_{i}')
    with c3:
        idx2 = STATUSES.index(ov['status']) if ov['status'] in STATUSES else 0
        st.session_state['overrides'][i]['status'] = st.selectbox(
            'Status', STATUSES, index=idx2, key=f'ov_stat_{i}')
    with c4:
        st.write(''); st.write('')
        if st.button('✕', key=f'del_{i}') and len(st.session_state['overrides']) > 1:
            st.session_state['overrides'].pop(i)
            st.rerun()

if st.button('+ Add Row', key='add_ov'):
    st.session_state['overrides'].append({'e': '', 'cti': CTI_OFFICES[0], 'status': STATUSES[0]})
    st.rerun()

st.divider()

# Step 3
st.header('Step 3 — Run')

if joiner_file and zoho_file:
    if st.button('▶ Fill + Split by Office', type='primary', use_container_width=True):
        joiner_bytes   = joiner_file.read()
        zoho_bytes     = zoho_file.read()
        inactive_bytes = inactive_file.read() if inactive_file else None

        with st.spinner('Building lookup...'):
            lookup = build_lookup_dict(zoho_bytes, inactive_bytes)
        st.write(f'Lookup ready: **{len(lookup):,}** seafarers in master')

        manual = {
            ov['e'].strip(): {'cti': ov['cti'], 'status': ov['status']}
            for ov in st.session_state['overrides'] if ov['e'].strip()
        }
        if manual:
            st.write(f'Manual overrides: **{len(manual)}** entries')

        with st.spinner('Filling joiner report...'):
            filled_bytes, fill_log, missing = run_joiner_fill(joiner_bytes, lookup, manual)

        with st.spinner('Splitting by office...'):
            split_bytes, offices = run_split(filled_bytes)

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        st.session_state.update({
            'j_filled':      filled_bytes,
            'j_filled_name': f'Joiner_Filled_{ts}.xlsx',
            'j_split':       split_bytes,
            'j_split_name':  f'CTI_By_Office_{ts}.zip',
            'j_log':         fill_log,
            'j_missing':     missing,
            'j_offices':     offices,
        })
else:
    st.info('Upload the Joiner Report and Zoho Master to continue.')

# Results
if 'j_filled' in st.session_state:
    st.divider()
    st.header('Results')
    st.code(st.session_state['j_log'])

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            '⬇ Filled Joiner Report (.xlsx)',
            data=st.session_state['j_filled'],
            file_name=st.session_state['j_filled_name'],
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True, key='dl_filled')
    with c2:
        st.download_button(
            '⬇ Split by Office (.zip)',
            data=st.session_state['j_split'],
            file_name=st.session_state['j_split_name'],
            mime='application/zip',
            use_container_width=True, key='dl_split')

    if st.session_state.get('j_offices'):
        st.info('Offices generated: ' + ' · '.join(st.session_state['j_offices']))

    if st.session_state.get('j_missing'):
        missing = st.session_state['j_missing']
        st.warning(f'⚠ {len(missing)} rows not found in master. Add their E-Numbers in Step 2 and re-run.')
        st.dataframe(pd.DataFrame(missing), use_container_width=True, hide_index=True)

st.divider()
st.caption('CTI Indonesia · Bali · Internal Use Only')
