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

def frac(c,k,Xo): return 1.0/(1.0+math.exp(k*(c+Xo)))
def phev_w(c): return 0.0 + 0.25*frac(c,-0.16,-20)  # PHEV stays on shared block, untouched

def othersum_at(i,TTLE):
    c = counts[i]
    s = 0.0
    for tech in ["NG","gas","diesel","LPG","H2"]:
        s += TTS_fixed[tech]*math.exp(TTLE*cost[tech][i])
    s += phev_w(c)*math.exp(TTLE*cost["PHEV"][i])
    return s

def evaluate(TTLE,k,Xo):
    fr = [frac(c,k,Xo) for c in counts]
    W_needed = []
    for i,y in enumerate(years):
        os_ = othersum_at(i,TTLE)
        s = target_share[i]
        W_needed.append((s/(1.0-s)) * os_ * math.exp(-TTLE*cost["BEV"][i]))
    A11 = sum((1-f)**2 for f in fr); A12 = sum((1-f)*f for f in fr); A22 = sum(f*f for f in fr)
    b1 = sum((1-f)*w for f,w in zip(fr,W_needed)); b2 = sum(f*w for f,w in zip(fr,W_needed))
    det = A11*A22-A12*A12
    D = (A22*b1-A12*b2)/det; E=(A11*b2-A12*b1)/det
    print(f"TTLE={TTLE} k={k} Xo={Xo} -> D={D:.4f} E={E:.4f}")
    for i,y in enumerate(years):
        f = fr[i]
        W = D+(E-D)*f
        os_ = othersum_at(i,TTLE)
        share = W*math.exp(TTLE*cost["BEV"][i])/(W*math.exp(TTLE*cost["BEV"][i])+os_)
        print(f"  {y}: target={target_share[i]*100:.2f}% fitted={share*100:.2f}% diff={100*(share-target_share[i]):+.2f}pp")
    return D,E

print("Candidate A: TTLE=-0.6, k=-0.16 (unchanged), Xo=-38")
evaluate(-0.6,-0.16,-38)
