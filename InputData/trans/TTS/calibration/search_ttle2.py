import math

cost = {
    "BEV":     [4.5374, 3.19655, 2.9809,  2.91218, 2.87752, 2.81443],
    "NG":      [3.59233,3.72923, 3.80016, 3.851,   3.88124, 3.90367],
    "gas":     [3.20882,3.22052, 3.23184, 3.24144, 3.25166, 3.26537],
    "diesel":  [3.7685, 3.81035, 3.8517,  3.87094, 3.89173, 3.9149],
    "PHEV":    [4.23513,3.47043, 3.45192, 3.44252, 3.44528, 3.42325],
    "LPG":     [3.62255,3.6961,  3.72534, 3.75866, 3.78749, 3.81516],
    "H2":      [13.9549,7.96097, 7.37328, 6.97498, 6.6695,  6.43188],
}
years = [2024,2030,2035,2040,2045,2050]
counts = [y - 2019 for y in years]
L,k,Xo = 1.0, -0.16, -20.0
def frac(c):
    return 1.0/(1.0+math.exp(k*(c+Xo)))
TTS_fixed = {"NG":0.0006, "gas":0.01, "diesel":1.0, "LPG":0.1, "H2":0.0005}
PHEV_D, PHEV_E = 0.0, 0.25
target_share = [0.082, 0.1718, 0.2614, 0.3671, 0.4743, 0.6285]

def othersum_at(i,TTLE):
    c = counts[i]
    s = 0.0
    for tech in ["NG","gas","diesel","LPG","H2"]:
        s += TTS_fixed[tech]*math.exp(TTLE*cost[tech][i])
    phev_w = PHEV_D + (PHEV_E-PHEV_D)*frac(c)
    s += phev_w*math.exp(TTLE*cost["PHEV"][i])
    return s

def fit_and_eval(TTLE):
    W_needed = []
    for i,y in enumerate(years):
        os_ = othersum_at(i,TTLE)
        s = target_share[i]
        W = (s/(1.0-s)) * os_ * math.exp(-TTLE*cost["BEV"][i])
        W_needed.append(W)
    fr = [frac(c) for c in counts]
    A11 = sum((1-f)**2 for f in fr); A12 = sum((1-f)*f for f in fr); A22 = sum(f*f for f in fr)
    b1 = sum((1-f)*w for f,w in zip(fr,W_needed)); b2 = sum(f*w for f,w in zip(fr,W_needed))
    det = A11*A22-A12*A12
    D = (A22*b1-A12*b2)/det; E=(A11*b2-A12*b1)/det
    maxerr = 0.0
    errs=[]
    for i,y in enumerate(years):
        f = fr[i]
        W = D+(E-D)*f
        os_ = othersum_at(i,TTLE)
        share = W*math.exp(TTLE*cost["BEV"][i])/(W*math.exp(TTLE*cost["BEV"][i])+os_)
        err = 100*(share-target_share[i])
        errs.append(err)
        maxerr = max(maxerr, abs(err))
    neg = D<0 or E<0
    return D,E,errs,maxerr,neg,W_needed

print("W_needed at TTLE=-8 (sanity check on severity):")
_,_,_,_,_,W8 = fit_and_eval(-8)
print(W8)

print("\nSearching TTLE:")
for TTLE in [-1.3,-1.25,-1.2,-1.15,-1.1,-1.05,-1.0,-0.98,-0.96,-0.94,-0.92,-0.9]:
    D,E,errs,maxerr,neg,W = fit_and_eval(TTLE)
    ratio = max(W)/min(W) if min(W)>0 else float('inf')
    monotonic = all(W[i] <= W[i+1]*1.001 for i in range(len(W)-1))
    print(f"TTLE={TTLE:+.2f} D={D:.4f} E={E:.4f} neg={neg} max/min_W={ratio:8.2f} monotonic={monotonic} errs(pp)={['%.2f'%e for e in errs]} max|err|={maxerr:.2f}")
