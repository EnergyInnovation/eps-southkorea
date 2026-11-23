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
TTS_fixed = {"NG":0.0006, "gas":0.01, "diesel":1.0, "LPG":0.1, "H2":0.0005}
target_share = [0.082, 0.1718, 0.2614, 0.3671, 0.4743, 0.6285]

def frac(c,k,Xo):
    return 1.0/(1.0+math.exp(k*(c+Xo)))

def phev_w(c,k,Xo,PHEV_D=0.0,PHEV_E=0.25):
    return PHEV_D + (PHEV_E-PHEV_D)*frac(c,k,Xo)

def othersum_at(i,TTLE,k,Xo,phev_uses_shared=True,phev_k=-0.16,phev_Xo=-20):
    c = counts[i]
    s = 0.0
    for tech in ["NG","gas","diesel","LPG","H2"]:
        s += TTS_fixed[tech]*math.exp(TTLE*cost[tech][i])
    pk,pXo = (k,Xo) if phev_uses_shared else (phev_k,phev_Xo)
    s += phev_w(c,pk,pXo)*math.exp(TTLE*cost["PHEV"][i])
    return s

def fit_and_eval(TTLE,k,Xo):
    fr = [frac(c,k,Xo) for c in counts]
    W_needed = []
    for i,y in enumerate(years):
        os_ = othersum_at(i,TTLE,k,Xo)
        s = target_share[i]
        W = (s/(1.0-s)) * os_ * math.exp(-TTLE*cost["BEV"][i])
        W_needed.append(W)
    A11 = sum((1-f)**2 for f in fr); A12 = sum((1-f)*f for f in fr); A22 = sum(f*f for f in fr)
    b1 = sum((1-f)*w for f,w in zip(fr,W_needed)); b2 = sum(f*w for f,w in zip(fr,W_needed))
    det = A11*A22-A12*A12
    if abs(det) < 1e-12:
        return None
    D = (A22*b1-A12*b2)/det; E=(A11*b2-A12*b1)/det
    maxerr = 0.0
    for i,y in enumerate(years):
        f = fr[i]
        W = D+(E-D)*f
        os_ = othersum_at(i,TTLE,k,Xo)
        share = W*math.exp(TTLE*cost["BEV"][i])/(W*math.exp(TTLE*cost["BEV"][i])+os_)
        err = 100*(share-target_share[i])
        maxerr = max(maxerr, abs(err))
    neg = D<0 or E<0
    return D,E,maxerr,neg

best = None
for TTLE in [-0.8,-0.75,-0.7,-0.65,-0.6,-0.55,-0.5,-0.45,-0.4]:
    for k in [-0.22,-0.2,-0.18,-0.16,-0.14,-0.12,-0.1]:
        for Xo10 in range(-450,-300,2):
            Xo = Xo10/10.0
            r = fit_and_eval(TTLE,k,Xo)
            if r is None: continue
            D,E,maxerr,neg = r
            if neg: continue
            if best is None or maxerr < best[0]:
                best = (maxerr, TTLE,k,Xo,D,E)

print("Best (dedicated shape block) fit found:")
print(f"maxerr={best[0]:.2f}pp TTLE={best[1]} k={best[2]} Xo={best[3]} D={best[4]:.4f} E={best[5]:.4f}")

# also show what happens keeping k,Xo at the shared value but scanning TTLE finely (baseline for comparison)
print("\nFor comparison, best with FIXED shared k=-0.16,Xo=-20 (already found ~8.2pp):")
r = fit_and_eval(-1.0,-0.16,-20)
print(r)
