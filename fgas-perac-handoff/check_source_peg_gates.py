"""Gates for step 2 (per-source F-gas peg) of the consolidated South Korea F-gas handoff.

  python check_source_peg_gates.py check      # after Apply-SourcePeg.ps1 + Export-PeracFgas.ps1 -Mode Export

Take a baseline first with:  python check_fgas_gates.py snapshot   (do it BEFORE step 2 edits).

Gate 1 target   : PERAC's pasted multiplier block equals the BPE tab; the BPE tab equals an independent
                  recomputation (BAU group(t) / EPA baseline group(t)); each CSV's summed potential per anchor year
                  equals sum over EPA_MACCs rows (KOR, that policy) of quantity x multiplier(year). HFC-23 (HCFC) contributes 0.
                  Tolerance 1e-5 relative (CSVs carry 6 significant digits). Then the offline sweep: a $10/t
                  (2012 USD) tax must no longer abate 100% of F-gases in 2025-2035.
Gate 2 control  : only the workbook and the four step-2 CSVs changed since the snapshot; the six-file change
                  set against HEAD is reported (workbook, five PERAC CSVs, WebAppData.xlsx).
Gate 3 invariant: total F-gas potential across the four CSVs <= BPE-F-gases chemicals 20 in every year
                  2021-2050; alum-IMR <= BPE other metals 242; no negatives; 2020 column zero; zero CRLF.
Needs openpyxl and git on PATH.
"""
import csv, io, os, re, subprocess, sys, tempfile
import openpyxl
import check_fgas_gates as base

REPO = base.REPO
PERAC = base.PERAC
BOOK = base.BOOK
HERE = os.path.dirname(os.path.abspath(__file__))
POL = {  # csv name -> EPA_MACCs "assigned policy"
    'fgassubstitution': 'F-gas substitution', 'fgasdestruction': 'F-gas destruction',
    'fgasrecovery': 'F-gas recovery and recycling',
    'chem-inspctmaintretrofit': 'chemicals - F-gas inspection maintenance retrofitting'}
GROUPS = {  # EPA_MACCs source code -> (EPA baseline source name, group, inventory industry row or None)
    'SEMI': ('Semiconductor Manufacturing', 'elec', 'computers and electronics 26'),
    'FPD': ('Flat Panel Display Manufacturing', 'elec', 'computers and electronics 26'),
    'PV': ('Photovoltaic Cell Manufacturing', 'elec', 'computers and electronics 26'),
    'RAC': ('Refrigeration and Air Conditioning', 'ods', 'appliances and electrical equipment 27'),
    'FOAM': ('Foams Mfg, Use, and Disposal', 'ods', 'appliances and electrical equipment 27'),
    'AERO': ('Aerosols', 'ods', 'appliances and electrical equipment 27'),
    'FIRE': ('Fire Protection', 'ods', 'appliances and electrical equipment 27'),
    'SOLV': ('Solvent Use', 'ods', 'appliances and electrical equipment 27'),
    'EPS': ('Electric Power Systems', 'sf6', 'other manufacturing 31T33'),
    'HCFC': ('HCFC-22 Production', 'hfc23', None)}
YEARS = (2025, 2030, 2035, 2040, 2045, 2050)

def multipliers():
    """Year-varying multiplier per EPA source code, recomputed independently:
    BAU F-gases of the group (BPE-F-gases_pre industry row, read from the BPE workbook) / EPA Korea baseline of
    the group's sources (data/epa_korea_fgas_baseline_2020_2050.csv)."""
    bw = openpyxl.load_workbook(os.path.join(REPO, 'InputData/indst/BPE/BAU Process Emissions.xlsx'), read_only=True, data_only=True)
    pre = list(bw['BPE-F-gases_pre'].iter_rows(min_row=2, max_row=45, max_col=32, values_only=True))
    years = [int(y) for y in pre[0][1:32]]
    bau = {}
    for r in pre[1:]:
        if r[0]: bau[r[0]] = {y: (r[1 + k] or 0) / 1e12 for k, y in enumerate(years)}
    epa = {}
    with open(os.path.join(HERE, 'data', 'epa_korea_fgas_baseline_2020_2050.csv'), newline='', encoding='utf-8') as f:
        for r in csv.DictReader(f): epa[(r['source'], int(r['year']))] = epa.get((r['source'], int(r['year'])), 0) + float(r['MMT CO2e'])
    out = {}
    for code, (name, grp, ind) in GROUPS.items():
        codes = [c for c, v in GROUPS.items() if v[1] == grp]
        out[code] = {}
        for y in years:
            if ind is None: out[code][y] = 0.0; continue
            den = sum(epa.get((GROUPS[c][0], y), 0) for c in codes)
            out[code][y] = bau[ind][y] / den if den else 0.0
    return out

def csv_totals(name, text=None):
    text = text or open(os.path.join(REPO, f'{PERAC}PERAC-{name}.csv'), newline='').read().replace('\r\n', '\n')
    rows = list(csv.reader(io.StringIO(text))); hdr = rows[1]
    return {int(y): sum(float(r[j] or 0) for r in rows[2:]) / 1e12 for j, y in enumerate(hdr) if j >= 2}

def gate1():
    ok = True
    wb = openpyxl.load_workbook(os.path.join(REPO, BOOK), read_only=True, data_only=True)
    mult = multipliers()
    print('  multipliers (recomputed here from BPE-F-gases_pre and data/*.csv), 2021 / 2035 / 2050:',
          {c: [round(mult[c][y], 3) for y in (2021, 2035, 2050)] for c in ('SEMI', 'RAC', 'EPS', 'HCFC')})
    # (a) the block pasted into PERAC equals the BPE tab, every year
    bpe = openpyxl.load_workbook(os.path.join(REPO, 'InputData/indst/BPE/BAU Process Emissions.xlsx'), read_only=True, data_only=True)['F-gas Source Multipliers']
    bt = {r[0]: r[2:33] for r in bpe.iter_rows(min_row=16, max_row=19, max_col=33, values_only=True)}
    ma = wb['Multipliers and Adjustments']
    hdr = [c.value for c in next(ma.iter_rows(min_row=22, max_row=22, max_col=35))]
    worst = 0.0
    for r in ma.iter_rows(min_row=505, max_row=520, max_col=35, values_only=True):
        if r[2] in bt:
            for c in range(4, 35):
                y = int(hdr[c]); worst = max(worst, abs(r[c] - bt[r[2]][y - 2020]))
    print(f'  PERAC multiplier block vs BPE tab: max abs diff {worst:.1e}'); ok &= worst < 1e-12
    # (b) BPE tab vs the independent recomputation
    gname = {'SEMI': 'Electronics', 'RAC': 'ODS substitutes', 'EPS': 'Electrical equipment SF6'}
    worst = max(abs(bt[gname[c]][y - 2020] - mult[c][y]) for c in gname for y in range(2020, 2051))
    print(f'  BPE tab vs independent recomputation: max abs diff {worst:.1e}'); ok &= worst < 1e-9
    # (c) CSV totals vs EPA_MACCs raw x multiplier(year)
    exp = {n: {y: 0.0 for y in YEARS} for n in POL}
    for r in wb['EPA_MACCs'].iter_rows(min_row=2, values_only=True):
        if r[3] != 'KOR' or r[2] not in YEARS: continue
        for n, p in POL.items():
            if r[7] == p: exp[n][r[2]] += r[6] * mult[r[1]][r[2]]
    for n in POL:
        got = csv_totals(n)
        worst = max(abs(got[y] - exp[n][y]) / exp[n][y] for y in YEARS if exp[n][y] > 0)
        print(f'  {n}: max rel diff vs recomputed {worst:.2e}; 2035 potential {got[2035]:.2f} Mt (expected {exp[n][2035]:.2f})')
        ok &= worst < 1e-5
    return ok

def sweep():
    r = subprocess.run([sys.executable, os.path.join(HERE, 'sweep_fgas_carbon_tax.py'),
                        os.path.join(tempfile.gettempdir(), 'fgas_sweep.json')], cwd=REPO, capture_output=True, text=True)
    if r.returncode: print(r.stderr[-400:]); return False
    sec = r.stdout.split('AS SHIPPED')[1].split('CORRECTED')[0]
    line = [l for l in sec.splitlines() if l.strip().startswith('lever   13.0')][0]
    print('  $10 tax:', line.strip())
    pc = {int(k): int(v) for k, v in re.findall(r'(\d{4}):\s*(\d+)%', line)}
    return all(pc[y] < 100 for y in (2025, 2030, 2035))

def git(*a): return subprocess.run(['git', *a], cwd=REPO, capture_output=True, text=True, check=True).stdout

def gate2():
    import json
    if not os.path.exists(base.SNAP): print('  no baseline snapshot'); return False
    old = json.load(open(base.SNAP)); now = base.hashes()
    changed = {p for p in set(old) | set(now) if old.get(p) != now.get(p)}
    allowed = {BOOK, 'InputData/indst/BPE/BAU Process Emissions.xlsx'} | {f'{PERAC}PERAC-{n}.csv' for n in POL}
    extra = sorted(changed - allowed); missing = sorted(allowed - changed)
    print('  changed since snapshot:', sorted(changed))
    if extra: print('  UNEXPECTED:', extra)
    if missing: print('  expected but unchanged:', missing)
    vs_head = sorted(l[3:].strip('"') for l in git('status', '--short', '--', 'InputData', 'WebAppData.xlsx').splitlines())
    six = {BOOK, 'WebAppData.xlsx'} | {f'{PERAC}PERAC-{n}.csv' for n in list(POL) + ['alum-inspctmaintretrofit']}
    print('  modified vs HEAD, not part of this change (pre-existing, other sessions):', [p for p in vs_head if p not in six])
    print('  six-file set missing from git status:', sorted(six - set(vs_head)))
    return not extra and not missing and not (six - set(vs_head))

def gate3():
    ok = True
    bpe = list(csv.reader(open(os.path.join(REPO, 'InputData/indst/BPE/BPE-F-gases.csv'), newline='')))
    bh = bpe[1]
    def bperow(ind): return {int(y): float(v) / 1e12 for y, v in zip(bh[1:], next(r for r in bpe[2:] if r[0] == ind)[1:]) if y}  # g -> Mt
    chem, mg = bperow('chemicals 20'), bperow('other metals 242')
    tot = {y: 0.0 for y in range(2021, 2051)}
    for n in list(POL) + ['alum-inspctmaintretrofit']:
        b = open(os.path.join(REPO, f'{PERAC}PERAC-{n}.csv'), 'rb').read()
        rows = list(csv.reader(io.StringIO(b.decode().replace('\r\n', '\n'))))
        bad = []
        if b.count(b'\r\n'): bad.append('CRLF')
        if len(rows) != 74 or len(rows[2]) != 33: bad.append('shape')
        if any(float(r[2] or 0) != 0 for r in rows[2:]): bad.append('2020 nonzero')
        if any(float(x or 0) < 0 for r in rows[2:] for x in r[2:]): bad.append('negative')
        print(f'  {n}: {"ok" if not bad else bad}'); ok &= not bad
        t = csv_totals(n)
        if n.startswith('alum'):
            over = [y for y in range(2021, 2051) if t[y] > mg[y]]
            print(f'  alum-IMR <= BPE other metals 242 in every year: {"yes" if not over else "NO in " + str(over)}'); ok &= not over
        else:
            for y in tot: tot[y] += t[y]
    over = [(y, round(tot[y], 1), round(chem[y], 1)) for y in tot if tot[y] > chem[y]]
    print(f'  total F-gas potential <= BPE chemicals 20 in every year 2021-2050: {"yes" if not over else "NO"}')
    for y in (2025, 2030, 2035, 2040, 2050): print(f'    {y}: potential {tot[y]:.1f} Mt vs BPE {chem[y]:.1f} Mt')
    if over: print('    years over:', over)
    return ok and not over

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'check':
        res = {}
        print('Gate 1 target'); res['1'] = gate1() and sweep()
        print('Gate 2 no-change control'); res['2'] = gate2()
        print('Gate 3 invariants'); res['3'] = gate3()
        print('\n'.join(f'Gate {k}: {"PASS" if v else "FAIL"}' for k, v in res.items()))
        sys.exit(0 if all(res.values()) else 1)
    sys.exit(__doc__)
