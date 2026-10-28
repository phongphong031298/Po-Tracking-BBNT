"""
data_manager.py
Quản lý việc đọc/ghi file Excel theo dõi PO, chuẩn hóa cấu trúc.
- Chỉ đọc các sheet ĐANG HIỂN THỊ (bỏ qua sheet ẩn/hidden).
- Chuẩn hóa tên cột về schema chung, bỏ cột Column1/Column2.
- Ghi đợt giao hàng mới vào đúng vị trí, có backup tự động.
"""
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from pathlib import Path
import shutil
from datetime import datetime


# =========================================================
# SCHEMA CHUẨN
# =========================================================
COMMON_COLS = [
    "Phân loại",                       # SI / -
    "Mã SP",
    "Tên SP (PO)",
    "Ngày giao hàng",
    "Số PO",
    "Số lượng trên ĐH",
    "Số lượng giao",
    "Chênh lệch (thừa so với ĐH)",     # hoặc "Còn lại so với ĐH"
    "Đơn giá",
    "Hoàn thành",
    "Ghi chú",
    "Dung sai",
]

# Map tên cột trong file gốc -> tên chuẩn
COL_ALIASES = {
    "Phân loại": ["SI", "Phân loại", "STT"],
    "Mã SP": ["Mã SP", "Mã sản phẩm", "Mã hàng"],
    "Tên SP (PO)": ["Tên SP\n(PO)", "Tên SP (PO)", "Tên SP", "Tên hàng"],
    "Ngày giao hàng": ["Ngày giao hàng", "Ngày"],
    "Số PO": ["Số PO", "Số đơn hàng", "PO"],
    "Số lượng trên ĐH": ["Số  lượng trên ĐH", "Số lượng trên ĐH", "SL trên ĐH"],
    "Số lượng giao": ["Số  lượng giao", "Số lượng giao", "SL giao"],
    "Chênh lệch (thừa so với ĐH)": [
        "Chênh lệch \n(thừa so với ĐH)",
        "Chênh lệch (thừa so với ĐH)",
        "Còn lại so với ĐH",
    ],
    "Đơn giá": ["Đơn giá"],
    "Hoàn thành": ["Hoàn thành", "Trạng thái"],
    "Ghi chú": ["Ghi chú"],
    "Dung sai": ["Dung sai"],
}


# =========================================================
# ĐỌC SHEET VISIBLE
# =========================================================
def get_visible_sheets(file_path: str) -> list[str]:
    """
    Trả về danh sách tên các sheet ĐANG HIỂN THỊ (không bị hide).
    sheet_state có thể là: 'visible', 'hidden', 'veryHidden'.
    """
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    visible = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
    wb.close()
    return visible


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Đổi tên cột về schema chuẩn, thêm cột thiếu, bỏ cột lạ."""
    rename_map = {}
    for std_name, aliases in COL_ALIASES.items():
        for col in df.columns:
            if str(col).strip() in [a.strip() for a in aliases]:
                rename_map[col] = std_name
                break
    df = df.rename(columns=rename_map)

    # Đảm bảo đủ cột
    for c in COMMON_COLS:
        if c not in df.columns:
            df[c] = None
    return df[COMMON_COLS]


def _find_header_row(raw: pd.DataFrame) -> int | None:
    """Tìm dòng chứa 'Mã SP' hoặc 'Mã sản phẩm' trong 15 dòng đầu."""
    for i in range(min(15, len(raw))):
        row_vals = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[i].tolist()]
        if any("Mã SP" in v or "Mã sản phẩm" in v for v in row_vals):
            return i
    return None


def load_all_sheets(file_path: str) -> dict[str, pd.DataFrame]:
    """
    Đọc tất cả các sheet ĐANG HIỂN THỊ trong file Excel.
    Bỏ qua sheet ẩn (hidden / veryHidden).
    """
    visible_sheets = get_visible_sheets(file_path)
    result = {}

    for sheet_name in visible_sheets:
        raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        header_row = _find_header_row(raw)

        if header_row is None:
            # Sheet không có cấu trúc PO (vd: Kiên Giang - tồn kho)
            result[sheet_name] = raw
            continue

        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
        df = _normalize_columns(df)
        df = df.dropna(how="all", subset=["Mã SP", "Số PO"])
        df["_sheet"] = sheet_name
        result[sheet_name] = df

    return result


# =========================================================
# DANH SÁCH SẢN PHẨM (autocomplete)
# =========================================================
def load_products_index(file_path: str) -> pd.DataFrame:
    """
    Trả về danh sách sản phẩm duy nhất (Mã SP, Tên SP) trong các sheet visible.
    Dùng cho autocomplete khi nhập đơn mới.
    """
    sheets = load_all_sheets(file_path)
    frames = []
    for _, df in sheets.items():
        if "Mã SP" not in df.columns:
            continue
        sub = df[["Mã SP", "Tên SP (PO)"]].copy()
        sub = sub.dropna(subset=["Mã SP"])
        sub["Mã SP"] = sub["Mã SP"].astype(str).str.strip()
        frames.append(sub)

    if not frames:
        return pd.DataFrame(columns=["Mã SP", "Tên SP (PO)"])

    all_products = pd.concat(frames, ignore_index=True)
    all_products = all_products.drop_duplicates(subset=["Mã SP"], keep="last")
    all_products = all_products.sort_values("Mã SP").reset_index(drop=True)
    return all_products


# =========================================================
# TÌM KIẾM
# =========================================================
def search_orders(
    file_path: str, keyword: str, sheet_filter: str | None = None
) -> pd.DataFrame:
    """Tìm kiếm theo Mã SP, Tên SP, hoặc Số PO (chỉ trên sheet visible)."""
    sheets = load_all_sheets(file_path)
    frames = []
    for name, df in sheets.items():
        if "Mã SP" not in df.columns:
            continue
        if sheet_filter and name != sheet_filter:
            continue
        mask = (
            df["Mã SP"].astype(str).str.contains(keyword, case=False, na=False)
            | df["Tên SP (PO)"].astype(str).str.contains(keyword, case=False, na=False)
            | df["Số PO"].astype(str).str.contains(keyword, case=False, na=False)
        )
        sub = df[mask].copy()
        if not sub.empty:
            sub["_sheet"] = name
            frames.append(sub)

    if not frames:
        return pd.DataFrame(columns=COMMON_COLS + ["_sheet"])
    return pd.concat(frames, ignore_index=True)


# =========================================================
# TỔNG HỢP THEO PO
# =========================================================
def get_po_summary(file_path: str) -> pd.DataFrame:
    """
    Tổng hợp theo từng PO: tổng SL đặt, tổng SL đã giao, còn lại.
    """
    sheets = load_all_sheets(file_path)
    rows = []
    for _, df in sheets.items():
        if "Số PO" not in df.columns:
            continue
        po_rows = df[df["Số PO"].notna() & (df["Số PO"].astype(str).str.strip() != "")]

        grouped = (
            po_rows.groupby("Số PO", dropna=True)
            .agg(
                MaSP=("Mã SP", "first"),
                TenSP=("Tên SP (PO)", "first"),
                SLDat=("Số lượng trên ĐH", "first"),
                SLGiao=("Số lượng giao", "sum"),
                DonGia=("Đơn giá", "first"),
                TrangThai=("Hoàn thành", "first"),
                Sheet=("_sheet", "first"),
            )
            .reset_index()
        )
        grouped["ConLai"] = grouped["SLDat"] - grouped["SLGiao"]
        rows.append(grouped)

    if not rows:
        return pd.DataFrame()
    return pd.concat(rows, ignore_index=True)


# =========================================================
# GHI ĐỢT GIAO HÀNG
# =========================================================
def append_delivery(
    file_path: str,
    sheet_name: str,
    ma_sp: str,
    so_po: str,
    ngay_giao: datetime,
    so_luong_giao: float,
    don_gia: float | None = None,
    ghi_chu: str = "",
    sl_tren_dh: float | None = None,  # Tham số mới
    ten_sp: str | None = None,        # Tham số mới
) -> int:
    """
    Ghi thêm 1 đợt giao hàng vào sheet tương ứng.
    - Nếu có sl_tren_dh, nghĩa là tạo PO mới -> Phân loại là "-"
    - Nếu không, nghĩa là giao hàng -> Phân loại là "" (trống)
    - Nếu có ten_sp, nghĩa là tạo SP mới -> Phân loại là "SI"
    """
    visible = get_visible_sheets(file_path)
    if sheet_name not in visible:
        raise ValueError(f"Sheet '{sheet_name}' đang bị ẩn hoặc không tồn tại.")

    backup = Path(file_path).with_suffix(".bak.xlsx")
    shutil.copy(file_path, backup)

    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]

    header_row = None
    for i in range(1, 20):
        for j in range(1, ws.max_column + 1):
            val = ws.cell(row=i, column=j).value
            if val and "Số PO" in str(val):
                header_row = i
                break
        if header_row:
            break

    if header_row is None:
        raise ValueError(f"Không tìm thấy cột 'Số PO' trong sheet '{sheet_name}'")

    col_map = {}
    for j in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=j).value
        if v:
            col_map[str(v).strip()] = j

    # Ánh xạ tên cột linh hoạt
    col_phan_loai = col_map.get("SI") or col_map.get("Phân loại")
    col_ma_sp = col_map.get("Mã SP")
    col_ten_sp = col_map.get("Tên SP\n(PO)") or col_map.get("Tên SP (PO)")
    col_po = col_map.get("Số PO")
    
    # Tìm các cột số lượng
    col_sl_dh = None
    for k in ["Số  lượng trên ĐH", "Số lượng trên ĐH"]:
        if k in col_map: col_sl_dh = col_map[k]; break
        
    col_sl_giao = None
    for k in ["Số  lượng giao", "Số lượng giao"]:
        if k in col_map: col_sl_giao = col_map[k]; break

    col_ngay = None
    for k in ["Ngày giao hàng", "Ngày"]:
        if k in col_map:
            col_ngay = col_map[k]
            break
    col_don_gia = col_map.get("Đơn giá")
    col_ghi_chu = col_map.get("Ghi chú")

    # Nếu tạo SP mới, thêm 1 dòng SI ở trên cùng
    if ten_sp:
        new_row = ws.max_row + 1
        if col_phan_loai: ws.cell(row=new_row, column=col_phan_loai).value = "SI"
        if col_ma_sp: ws.cell(row=new_row, column=col_ma_sp).value = ma_sp
        if col_ten_sp: ws.cell(row=new_row, column=col_ten_sp).value = ten_sp

    # Nếu tạo PO mới, thêm 1 dòng "-"
    if sl_tren_dh is not None:
        new_row = ws.max_row + 1
        if col_phan_loai: ws.cell(row=new_row, column=col_phan_loai).value = "-"
        if col_ma_sp: ws.cell(row=new_row, column=col_ma_sp).value = ma_sp
        if col_po: ws.cell(row=new_row, column=col_po).value = so_po
        if col_ngay: ws.cell(row=new_row, column=col_ngay).value = ngay_giao
        if col_sl_dh: ws.cell(row=new_row, column=col_sl_dh).value = sl_tren_dh
        if col_don_gia and don_gia: ws.cell(row=new_row, column=col_don_gia).value = don_gia

    # Thêm dòng giao hàng thực tế (Phân loại để trống)
    new_row = ws.max_row + 1
    if col_ma_sp: ws.cell(row=new_row, column=col_ma_sp).value = ma_sp
    if col_po: ws.cell(row=new_row, column=col_po).value = so_po
    if col_ngay: ws.cell(row=new_row, column=col_ngay).value = ngay_giao
    if col_sl_giao: ws.cell(row=new_row, column=col_sl_giao).value = so_luong_giao
    if col_ghi_chu: ws.cell(row=new_row, column=col_ghi_chu).value = ghi_chu

    wb.save(file_path)
    return new_row

def delete_po(file_path: str, sheet_name: str, ma_sp: str, so_po: str) -> int:
    """
    Xóa 1 PO: dòng thông tin đơn hàng (Phân loại "-") và toàn bộ các dòng
    giao hàng của PO đó (cùng Mã SP + Số PO).
    Chỉ XÓA NỘI DUNG các ô (giữ nguyên dòng) để không làm lệch các công thức
    =SUM(...) và vùng bảng Excel của những PO khác.
    Trả về số dòng đã xóa.
    """
    visible = get_visible_sheets(file_path)
    if sheet_name not in visible:
        raise ValueError(f"Sheet '{sheet_name}' đang bị ẩn hoặc không tồn tại.")

    backup = Path(file_path).with_suffix(".bak.xlsx")
    shutil.copy(file_path, backup)

    wb = openpyxl.load_workbook(file_path)  # KHÔNG dùng data_only để giữ công thức
    ws = wb[sheet_name]

    header_row = None
    for i in range(1, 20):
        for j in range(1, ws.max_column + 1):
            val = ws.cell(row=i, column=j).value
            if val and "Số PO" in str(val):
                header_row = i
                break
        if header_row:
            break
    if header_row is None:
        raise ValueError(f"Không tìm thấy cột 'Số PO' trong sheet '{sheet_name}'")

    col_ma_sp = col_po = None
    for j in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=j).value
        if not v:
            continue
        v = str(v).strip()
        if v == "Mã SP" and col_ma_sp is None:
            col_ma_sp = j
        if v == "Số PO" and col_po is None:
            col_po = j
    if col_ma_sp is None or col_po is None:
        raise ValueError("Không tìm thấy cột 'Mã SP' hoặc 'Số PO'")

    target_po = str(so_po).strip().upper()
    target_sp = str(ma_sp).strip().upper()

    count = 0
    for r in range(header_row + 1, ws.max_row + 1):
        po_val = ws.cell(row=r, column=col_po).value
        sp_val = ws.cell(row=r, column=col_ma_sp).value
        if po_val is None or sp_val is None:
            continue
        if str(po_val).strip().upper() == target_po and str(sp_val).strip().upper() == target_sp:
            for j in range(1, ws.max_column + 1):
                ws.cell(row=r, column=j).value = None
            count += 1

    if count == 0:
        raise ValueError(f"Không tìm thấy PO '{so_po}' của mã '{ma_sp}'.")

    wb.save(file_path)
    return count

def delete_row_by_index(file_path: str, sheet_name: str, row_index: int):
    """
    Xóa 1 dòng trong Excel theo index (index này là index thực tế của dòng trong file Excel, bắt đầu từ 1).
    Hàm này chỉ dùng cho tính năng xóa dòng trên App.
    """
    backup = Path(file_path).with_suffix(".bak.xlsx")
    shutil.copy(file_path, backup)

    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]
    
    # row_index truyền vào là index của dataframe (bắt đầu từ 0)
    # Trong openpyxl, dòng bắt đầu từ 1. Ta phải cộng thêm header_row (dòng tiêu đề) và 2 (vì index 0 -> dòng 1).
    # Nhưng vì mình tính toán dòng thực tế bên app.py rồi, nên ở đây chỉ cần truyền đúng dòng Excel là xong.
    
    ws.delete_rows(row_index, 1)
    wb.save(file_path)