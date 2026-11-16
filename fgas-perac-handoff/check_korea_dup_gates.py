"""Gates for the Korea EPA duplicate-copy correction (0.5 on every KOR quantity in EPA_MACCs, all blocks).

  python check_korea_dup_gates.py install   # after Export-PeracAll.ps1: copy staged CSVs whose values changed (>1e-5 rel)
  python check_korea_dup_gates.py check

Gate 1 target   : every PERAC CSV that reads EPA_MACCs equals 0.5 x its pre-correction version (cell by cell, tolerance
                  1e-5 relative, i.e. the 6-significant-digit export); F-gas totals equal EPA_MACCs raw x 0.5 x multiplier(year).
Gate 2 control  : PERAC-cement and PERAC-tillage (no EPA_MACCs input) and the four all-zero oil-and-gas tabs are
                  unchanged; the workbook's raw EPA_MACCs column G is unchanged vs the pre-correction backup.
Gate 3 invariant: total F-gas potential <= BPE chemicals 20 in every year 2021-2050; alum-IMR <= BPE other metals 242;
                  shape 74 x 33, 2020 column zero, no negatives, zero CRLF.
Needs openpyxl. Pre-correction backups: %TEMP%\\pre_dupfix_csv and %TEMP%\\perac_before_dupfix_backup.xlsx.
"""
import csv, io, os, shutil, sys
import openpyxl
import check_source_peg_gates as sp

REPO, PERAC, BOOK = sp.REPO, sp.PERAC, sp.BOOK
TMP = os.environ.get('TEMP') or '/tmp'
OLD = os.path.join(TMP, 'pre_dupfix_csv'); STAGE = os.path.join(TMP, 'peracall_stage')
NO_EPA = {'PERAC-cement.csv', 'PERAC-tillage.csv', 'PERAC-ngpProd-mthncptr.csv', 'PERAC-ngpProd-mthndstr.csv', 'PERAC-ngpPrcsTnD-mthndstr.csv'}
FACTOR = 0.5

def grid(path):
    return list(csv.reader(io.StringIO(open(path, 'rb').read().decode().replace('\r\n', '\n'))))[2:]

def maxrel(a, b, scale=1.0):
    w = 0.0
    for ra, rb in zip(a, b):
        for x, y in zip(ra[2:], rb[2:]):
            x = float(x or 0) * scale; y = float(y or 0)
            if max(abs(x), abs(y)) > 0: w = max(w, abs(x - y) / max(abs(x), abs(y)))
    return w

def install():
    for f in sorted(os.listdir(STAGE)):
        if maxrel(grid(os.path.join(OLD, f)), grid(os.path.join(STAGE, f))) > 1e-5:
            shutil.copyfile(os.path.join(STAGE, f), os.path.join(REPO, PERAC, f)); print('installed', f)
        else: print('unchanged (not copied)', f)

def gate1():
    ok = True
    for f in sorted(os.listdir(OLD)):
        if f in NO_EPA or f == 'PERAC-MCD.csv': continue
        w = maxrel(grid(os.path.join(OLD, f)), grid(os.path.join(REPO, PERAC, f)), FACTOR)
        print(f'  {f:36s} new vs 0.5 x old: max rel diff {w:.1e}'); ok &= w < 1e-5
    wb = openpyxl.load_workbook(os.path.join(REPO, BOOK), read_only=True, data_only=True)
    mult = sp.multipliers()
    exp = {n: {y: 0.0 for y in sp.YEARS} for n in sp.POL}
    for r in wb['EPA_MACCs'].iter_rows(min_row=2, values_only=True):
        if r[3] != 'KOR' or r[2] not in sp.YEARS: continue
        for n, p in sp.POL.items():
            if r[7] == p: exp[n][r[2]] += FACTOR * r[6] * mult[r[1]][r[2]]
    for n in sp.POL:
        got = sp.csv_totals(n)
        w = max(abs(got[y] - exp[n][y]) / exp[n][y] for y in sp.YEARS if exp[n][y] > 0)
        print(f'  {n}: F-gas total vs 0.5 x raw x multiplier: max rel diff {w:.1e}; 2035 potential {got[2035]:.2f} Mt'); ok &= w < 1e-5
    return ok

def gate2():
    ok = True
    for f in sorted(NO_EPA):
        same = open(os.path.join(OLD, f), 'rb').read() == open(os.path.join(REPO, PERAC, f), 'rb').read()
        print(f'  {f}: {"unchanged" if same else "CHANGED"}'); ok &= same
    bk = os.path.join(TMP, 'perac_before_dupfix_backup.xlsx')
    def raw(path):
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True)['EPA_MACCs']
        return [r[:7] for r in ws.iter_rows(min_row=2, values_only=True)]
    a, b = raw(bk), raw(os.path.join(REPO, BOOK))
    same = a == b
    print(f'  EPA_MACCs raw columns A:G identical to pre-correction backup ({len(a)} rows): {"yes" if same else "NO"}'); ok &= same
    return ok

def gate3():
    return sp.gate3()

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'install': install(); sys.exit(0)
    if len(sys.argv) > 1 and sys.argv[1] == 'check':
        res = {}
        print('Gate 1 target'); res['1'] = gate1()
        print('Gate 2 no-change control'); res['2'] = gate2()
        print('Gate 3 invariants'); res['3'] = gate3()
        print('\n'.join(f'Gate {k}: {"PASS" if v else "FAIL"}' for k, v in res.items()))
        sys.exit(0 if all(res.values()) else 1)
    sys.exit(__doc__)
