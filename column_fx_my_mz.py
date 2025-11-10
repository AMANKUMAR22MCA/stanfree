import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT
import openpyxl

# ============================================================
# SETTINGS
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# save output in same folder as this script:
OUTPUT_XLSX = os.path.join(BASE_DIR, "vertical_members_nodes_fx_myz.xlsx")

UNIT_TO_MM = 25.4
TOL = 1e-3

# ============================================================
# HELPERS
# ============================================================
def approx_eq(a, b, tol=TOL):
    return abs(a - b) <= tol

def dispid(obj, name):
    try:
        return obj._oleobj_.GetIDsOfNames(name)
    except Exception:
        return None

def safe_save(wb_obj, path):
    try:
        wb_obj.save(path)
    except PermissionError:
        alt = path.replace(".xlsx", f"_backup_{int(time.time())}.xlsx")
        print(f"⚠ File locked, saved as backup: {alt}")
        wb_obj.save(alt)

# ============================================================
# OPENSTAAD ATTACH
# ============================================================
print("\n🔗 Connecting to STAAD via OpenSTAAD...")
pythoncom.CoInitialize()
os_app = Dispatch("StaadPro.OpenSTAAD")
geom = os_app.Geometry
prop = getattr(os_app, "Property", None)
print("✅ Connected to STAAD model.")

# ============================================================
# GEOMETRY HELPERS
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
# STEP 1: FETCH ALL MEMBERS
# ============================================================
total_members = geom.GetMemberCount
print(f"📦 Total members found: {total_members}")

member_data = {}

for m in range(1, int(total_members) + 1):
    n1, n2 = get_member_incidence(m)
    if not n1 or not n2:
        continue
    x1, y1, z1 = get_node_coordinates_mm(n1)
    x2, y2, z2 = get_node_coordinates_mm(n2)
    if None in (x1, y1, z1, x2, y2, z2):
        continue
    member_data[m] = {
        "n1": n1, "n2": n2,
        "x1": x1, "y1": y1, "z1": z1,
        "x2": x2, "y2": y2, "z2": z2
    }

# ============================================================
# STEP 2: FILTER ONLY COLUMNS (vertical members)
# ============================================================
columns = {}
for mid, info in member_data.items():
    if approx_eq(info["x1"], info["x2"]) and approx_eq(info["z1"], info["z2"]) and not approx_eq(info["y1"], info["y2"]):
        columns[mid] = info

print(f"🧱 Total column members detected: {len(columns)}")

# ============================================================
# STEP 3: BUILD NODE-WISE MAPPING (only nodes part of columns)
# ============================================================
results = []
# We will also collect unique node rows (so each node appears once with its top/bottom column info)
node_map = {}

for col_id, info in columns.items():
    n1, n2 = info["n1"], info["n2"]
    y1, y2 = info["y1"], info["y2"]

    # Determine top & bottom node based on Y
    if y1 > y2:
        top_node, bottom_node = n1, n2
    else:
        top_node, bottom_node = n2, n1

    # For bottom node: record bottom column id
    if bottom_node not in node_map:
        node_map[bottom_node] = {"node_id": bottom_node, "top_column_id": 0, "bottom_column_id": col_id}
    else:
        node_map[bottom_node]["bottom_column_id"] = col_id

    # For top node: record top column id
    if top_node not in node_map:
        node_map[top_node] = {"node_id": top_node, "top_column_id": col_id, "bottom_column_id": 0}
    else:
        node_map[top_node]["top_column_id"] = col_id

# Convert node_map to results list (only nodes that are part of columns)
for nid, info in node_map.items():
    # Prefer showing whichever column this node belongs to (bottom first, else top)
    col_id = info["bottom_column_id"] if info["bottom_column_id"] != 0 else info["top_column_id"]

    results.append([
        # col_id,  # ✅ FIXED: use actual Column (Member) ID instead of 0
        info["node_id"],
        info["top_column_id"],
        info["bottom_column_id"]
    ])

# ============================================================
# STEP 4: WRITE TO EXCEL (save in same folder as script)
# ============================================================
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Column_Node_Map"

headers = [
     "Node ID", "Top Column ID", "Bottom Column ID"
]
ws.append(headers)

for row in results:
    ws.append(row)

safe_save(wb, OUTPUT_XLSX)
print(f"✅ Column-node mapping saved to: {OUTPUT_XLSX}")
