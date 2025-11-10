import os
import pythoncom
from win32com.client import Dispatch, VARIANT
import openpyxl

# ============================================================
# SETTINGS
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "Outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "footing_members.xlsx")

UNIT_TO_MM = 25.4
TOL = 1e-3

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def approx_eq(a, b, tol=TOL):
    return abs(a - b) <= tol

def dispid(obj, name):
    try:
        return obj._oleobj_.GetIDsOfNames(name)
    except Exception:
        return None

# ============================================================
# CONNECT TO STAAD
# ============================================================
print("🔗 Connecting to STAAD...")
pythoncom.CoInitialize()
os_app = Dispatch("StaadPro.OpenSTAAD")
geom = os_app.Geometry
print("✅ Connected to STAAD Model.")

# ============================================================
# GEOMETRY FETCH
# ============================================================
def get_member_incidence(member_no):
    try:
        dpid = geom._oleobj_.GetIDsOfNames("GetMemberIncidence")
        n1 = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        n2 = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        geom._oleobj_.InvokeTypes(
            dpid, 0, pythoncom.DISPATCH_METHOD, (pythoncom.VT_EMPTY, 0),
            ((pythoncom.VT_I4, 0),
             (pythoncom.VT_VARIANT, 1),
             (pythoncom.VT_VARIANT, 1)),
            int(member_no), n1, n2
        )
        return int(n1.value), int(n2.value)
    except Exception:
        return None, None

def get_node_coordinates_mm(node_id):
    try:
        dpid = geom._oleobj_.GetIDsOfNames("GetNodeCoordinates")
        x = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
        y = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
        z = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_R8, 0.0)
        geom._oleobj_.InvokeTypes(
            dpid, 0, pythoncom.DISPATCH_METHOD, (pythoncom.VT_EMPTY, 0),
            ((pythoncom.VT_I4, 0),
             (pythoncom.VT_VARIANT, 1),
             (pythoncom.VT_VARIANT, 1),
             (pythoncom.VT_VARIANT, 1)),
            int(node_id), x, y, z
        )
        return float(x.value) * UNIT_TO_MM, float(y.value) * UNIT_TO_MM, float(z.value) * UNIT_TO_MM
    except Exception:
        return None, None, None

# ============================================================
# PASS 1: Collect member and node data
# ============================================================
total_members = geom.GetMemberCount
print(f"📦 Total members found: {total_members}")

member_data = {}
node_coords = {}

for m in range(1, int(total_members) + 1):
    n1, n2 = get_member_incidence(m)
    if not n1 or not n2:
        continue
    x1, y1, z1 = get_node_coordinates_mm(n1)
    x2, y2, z2 = get_node_coordinates_mm(n2)
    member_data[m] = {"n1": n1, "n2": n2, "y1": y1, "y2": y2}
    node_coords[n1] = (x1, y1, z1)
    node_coords[n2] = (x2, y2, z2)

# ============================================================
# STEP 2: Identify Footing Members
# ============================================================
global_min_y = min(y for (_, y, _) in node_coords.values())
print(f"🌍 Global minimum Y (Footing Level): {global_min_y:.3f} mm")

rows = []
for mid, info in member_data.items():
    n1, n2, y1, y2 = info["n1"], info["n2"], info["y1"], info["y2"]
    footing_flag = "YES" if (approx_eq(y1, global_min_y) or approx_eq(y2, global_min_y)) else "NO"
    rows.append([mid, n1, n2, footing_flag])

# ============================================================
# WRITE TO EXCEL
# ============================================================
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Footing_Members"
ws.append(["Member ID", "Start Node ID", "End Node ID", "Footing Member (YES/NO)"])

for r in rows:
    ws.append(r)

wb.save(OUTPUT_XLSX)
print(f"📄 Footing member report saved: {OUTPUT_XLSX}")
