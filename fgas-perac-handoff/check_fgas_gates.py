"""Gates for the South Korea PERAC F-gas multiplier fix (HANDOFF.md). Run from anywhere.

  python check_fgas_gates.py snapshot   # BEFORE Export-PeracFgas.ps1 -Mode Apply: hash every file under InputData
  python check_fgas_gates.py check      # AFTER Apply: runs all three gates

Gate 1 target   : each new CSV == HEAD CSV x (new multiplier / old multiplier for that year), rel tol 1e-5
                  (the CSVs are written to 6 significant digits, so 1e-6 would false-fail), then the offline
                  sweep reproduces the expected abatement at a $10 tax.
Gate 2 control  : only the workbook and the five CSVs changed since the snapshot (other sessions' edits that
                  already existed at snapshot time are tolerated because they are in the baseline).
Gate 3 invariant: 2020 column all zeros, no negatives, zero CRLF, in the five CSVs.
Needs openpyxl and git on PATH.
"""
import csv, hashlib, io, json, os, re, subprocess, sys, tempfile
import openpyxl

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PERAC = 'InputData/indst/PERAC/'
BOOK = PERAC + 'Proc Emis Reductions and Costs.xlsx'
SNAP = os.path.join(tempfile.gettempdir(), 'fgas_gates_baseline.json')
# csv name -> (multiplier row, GHG Inventory peg cell, expected new multiplier)
FILES = {
    'fgassubstitution': (37, 'F26', 0.59125), 'fgasdestruction': (37, 'F26', 0.59125),
    'fgasrecovery': (37, 'F26', 0.59125), 'chem-inspctmaintretrofit': (37, 'F26', 0.59125),
    'alum-inspctmaintretrofit': (42, 'F31', 0.69486),
}
ALLOWED = {BOOK} | {f'{PERAC}PERAC-{n}.csv' for n in FILES}
TOL = 1e-5

def git(*a, binary=False):
    r = subprocess.run(['git', *a], cwd=REPO, capture_output=True, check=True)
    return r.stdout if binary else r.stdout.decode()

def hashes():
    out = {}
    for root, _, fs in os.walk(os.path.join(REPO, 'InputData')):
        for f in fs:
            p = os.path.join(root, f)
            out[os.path.relpath(p, REPO).replace(os.sep, '/')] = hashlib.sha1(open(p, 'rb').read()).hexdigest()
    return out

def mults(path, row):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)['Multipliers and Adjustments']
    yrs = [c.value for c in next(ws.iter_rows(min_row=22, max_row=22))][4:35]
    vals = [c.value for c in next(ws.iter_rows(min_row=row, max_row=row))][4:35]
    return dict(zip(yrs, vals))

def table(text):
    rows = list(csv.reader(io.StringIO(text)))
    return rows[:2], rows[2:]

def gate1():
    ok = True
    old_book = os.path.join(tempfile.gettempdir(), 'fgas_gates_old_book.xlsx')
    open(old_book, 'wb').write(git('show', f'HEAD:{BOOK}', binary=True))
    wbnew = openpyxl.load_workbook(os.path.join(REPO, BOOK), read_only=True, data_only=True)
    for n, (row, peg, exp) in FILES.items():
        new_m = wbnew['GHG Inventory'][peg].value
        if abs(new_m - exp) > 1e-5: print(f'  {n}: peg {peg}={new_m} != expected {exp}'); ok = False
        old_m = mults(old_book, row)
        cur_m = mults(os.path.join(REPO, BOOK), row)
        hdr_new, d_new = table(open(os.path.join(REPO, f'{PERAC}PERAC-{n}.csv'), newline='').read().replace('\r\n', '\n'))
        hdr_old, d_old = table(git('show', f'HEAD:{PERAC}PERAC-{n}.csv').replace('\r\n', '\n'))
        if hdr_new != hdr_old or len(d_new) != len(d_old): print(f'  {n}: header/row layout changed'); ok = False; continue
        worst = 0.0; bad = 0
        for ro, rn in zip(d_old, d_new):
            if ro[:2] != rn[:2]: bad += 1; continue
            for j, y in enumerate(hdr_old[1][2:], start=2):
                y = int(y); a = float(ro[j] or 0); b = float(rn[j] or 0)
                ratio = cur_m[y] / old_m[y] if n.startswith('alum') else new_m / old_m[y]
                want = a * ratio
                err = abs(b - want) / abs(want) if want else abs(b)
                worst = max(worst, err); bad += err > TOL
        print(f'  {n}: max rel err {worst:.2e}, cells over {TOL:g}: {bad}'); ok &= bad == 0
    return ok

def sweep():
    r = subprocess.run([sys.executable, os.path.join(REPO, 'fgas-perac-handoff', 'sweep_fgas_carbon_tax.py'),
                        os.path.join(tempfile.gettempdir(), 'fgas_sweep.json')], cwd=REPO, capture_output=True, text=True)
    if r.returncode: print(r.stderr[-400:]); return False
    sec = r.stdout.split('AS SHIPPED')[1] if 'AS SHIPPED' in r.stdout else r.stdout
    line = [l for l in sec.splitlines() if l.strip().startswith('lever   13.0')][0]
    print('  $10 tax, shipped-files section:', line.strip())
    pc = {int(k): int(v) for k, v in re.findall(r'(\d{4}):\s*(\d+)%', line)}
    return 56 <= pc[2025] <= 60 and 62 <= pc[2030] <= 66

def gate2():
    if not os.path.exists(SNAP): print('  no baseline: run "snapshot" before Apply'); return False
    base = json.load(open(SNAP)); now = hashes()
    changed = {p for p in set(base) | set(now) if base.get(p) != now.get(p)}
    extra = sorted(changed - ALLOWED); missing = sorted(ALLOWED - changed)
    print('  changed since snapshot:', sorted(changed))
    if extra: print('  UNEXPECTED changes:', extra)
    if missing: print('  expected but unchanged:', missing)
    return not extra and not missing

def gate3():
    ok = True
    for n in FILES:
        b = open(os.path.join(REPO, f'{PERAC}PERAC-{n}.csv'), 'rb').read()
        _, d = table(b.decode().replace('\r\n', '\n'))
        crlf = b.count(b'\r\n'); z = all(float(r[2] or 0) == 0 for r in d)
        neg = any(float(x or 0) < 0 for r in d for x in r[2:])
        print(f'  {n}: CRLF={crlf} 2020_all_zero={z} negatives={neg}'); ok &= crlf == 0 and z and not neg
    return ok

if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else ''
    if mode == 'snapshot':
        json.dump(hashes(), open(SNAP, 'w')); print('baseline saved:', SNAP)
    elif mode == 'check':
        res = {}
        print('Gate 1 target'); res['1'] = gate1() and sweep()
        print('Gate 2 no-change control'); res['2'] = gate2()
        print('Gate 3 invariants'); res['3'] = gate3()
        print('\n'.join(f'Gate {k}: {"PASS" if v else "FAIL"}' for k, v in res.items()))
        sys.exit(0 if all(res.values()) else 1)
    else:
        sys.exit(__doc__)
