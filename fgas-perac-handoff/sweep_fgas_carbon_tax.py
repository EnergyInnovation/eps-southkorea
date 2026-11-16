# Offline replica of the EPS carbon-tax -> PERAC process-emissions response for F-gases.
# Mirrors EPS.mdl: Process Emissions Abated due to Carbon Tax by Marginal Cost (bin abated in full if
# MCD>0 and tax>MCD), the per-measure cap MAX(dec, -BAU*share) and the MAX(...,0) floor. Ignores the
# production feedback (1+Percent Change in Production due to Policies).
import csv,sys,json,openpyxl
D='InputData/indst/'
def perac(name):
    rows=list(csv.reader(open(D+'PERAC/PERAC-'+name+'.csv')))
    h=rows[1]; return h,[(float(r[0]),float(r[1]),[float(x or 0) for x in r[2:]]) for r in rows[2:] if r and r[0]]
mcd=[float(r[1]) for r in csv.reader(open(D+'PERAC/PERAC-MCD.csv')) if r and r[0].startswith('PriceBin')]
bpe=list(csv.reader(open(D+'BPE/BPE-F-gases.csv'))); bh=bpe[1]
def bperow(ind):
    for r in bpe[2:]:
        if r[0]==ind: return {int(y):float(v) for y,v in zip(bh[1:],r[1:]) if y}
chem=bperow('chemicals 20')
# multipliers actually used (row 37) vs the GHG Inventory peg (F26)
wb=openpyxl.load_workbook(D+'PERAC/Proc Emis Reductions and Costs.xlsx',read_only=True,data_only=True)
ws=wb['Multipliers and Adjustments']
yrs=[c.value for c in next(ws.iter_rows(min_row=22,max_row=22))][4:35]
m37=[c.value for c in next(ws.iter_rows(min_row=37,max_row=37))][4:35]
mult=dict(zip(yrs,m37))
peg=wb['GHG Inventory']['F26'].value
print('GHG Inventory F26 peg =',peg)
meas=['fgassubstitution','fgasdestruction','fgasrecovery','chem-inspctmaintretrofit']
P={m:perac(m) for m in meas}
USD_PER_UNIT=0.771001481   # OCCF-DpSOCU: 2012 USD per 1000 KRW
def run(lever, years=(2025,2030,2035,2040,2050), fix=False):
    tax=lever*USD_PER_UNIT; out={}
    for y in years:
        tot=0; potall=0
        for m in meas:
            h,data=P[m]; j=h.index(str(y))-2
            sc=(peg/mult[y]) if fix else 1
            for i,(lo,hi,v) in enumerate(data):
                potall+=v[j]*sc
                if mcd[i]>0 and tax>mcd[i]: tot+=v[j]*sc
        b=chem[y]; out[y]=dict(bau=b/1e12,pot=potall/1e12,abated=min(tot,b)/1e12,pct=min(tot,b)/b)
    return out
levers=[0,1.3,5.2,6.5,7.8,10.2,13,26,65,130]
res={}
for fix in (False,True):
    print('\n=== ', 'CORRECTED (constant peg %.3f)'%peg if fix else 'AS SHIPPED (row 37 multipliers)')
    for L in levers:
        r=run(L,fix=fix); res[(fix,L)]=r
        print(f"lever {L:6.1f} (=${L*USD_PER_UNIT:6.1f}/t 2012USD): "+'  '.join(f"{y}:{r[y]['pct']:5.0%}" for y in r))
    r=run(0,fix=fix); print('  potential/BAU:', '  '.join(f"{y}:{r[y]['pot']:.0f}/{r[y]['bau']:.0f}Mt" for y in r))
json.dump({f"{'fix' if k[0] else 'asis'}_{k[1]}":v for k,v in res.items()},open(sys.argv[1],'w'))
