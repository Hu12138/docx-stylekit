TWIPS_PER_CM = 567.0
HALFPOINT_PER_PT = 2.0

def twips_to_cm(v):
    return round(float(v) / TWIPS_PER_CM, 2)

def cm_to_twips(cm):
    return int(round(float(cm) * TWIPS_PER_CM))

def pt_to_halfpoints(pt):
    return int(round(float(pt) * HALFPOINT_PER_PT))

def halfpoints_to_pt(hp):
    return float(hp) / HALFPOINT_PER_PT
