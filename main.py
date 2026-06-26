import flet as ft
import pandas as pd
import unicodedata
import re
import os
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

def normalize_text(text: str) -> str:
    if not text:
        return ""
    s = str(text).replace("Đ", "D").replace("đ", "d")
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    ).lower().strip()

def clean_unit_name(unit_str: str) -> str:
    if not unit_str or str(unit_str).strip().lower() in ['nan', '', '-', 'nat']:
        return "Chưa cập nhật"
    return str(unit_str).strip()

def clean_cell_to_date_str(val) -> str:
    if pd.isna(val) or str(val).strip().lower() in ['nan', '', 'nat', '-', '.', '0']:
        return ""
    if isinstance(val, (datetime, date, pd.Timestamp)):
        return val.strftime('%d/%m/%Y')

    raw_s = str(val).strip().split()[0]
    s_slashes = raw_s.replace('-', '/').replace('.', '/')

    for fmt in ('%d/%m/%Y', '%Y/%m/%d', '%m/%d/%Y', '%d/%m/%y', '%Y%m%d'):
        try:
            test_str = raw_s if fmt == '%Y%m%d' else s_slashes
            dt = datetime.strptime(test_str, fmt)
            return dt.strftime('%d/%m/%Y')
        except ValueError:
            continue
    return str(val).strip()

def find_column_by_keywords(columns, keywords) -> str | None:
    for kw in keywords:
        for col in columns:
            if kw == normalize_text(col):
                return col
    for kw in keywords:
        for col in columns:
            if kw in normalize_text(col):
                return col
    return None

class AllowanceDashboardApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Hệ thống Quản lý Thu hút & Phụ cấp"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.bgcolor = ft.Colors.BLUE_GREY_50
        self.page.padding = 20

        # Dữ liệu gốc và dữ liệu sau bộ lọc
        self.master_data = {}
        self.filtered_data = {}
        self.extra_headers = []

        # Các biến phục vụ phân trang
        self.per_page = 50          
        self.page_number = 1        
        self.stats_page_number = 1  
        self.adjust_page_number = 1
        self.stats_filtered_items = [] 

        # UI cho Tab 1: Danh Sách
        self.search_field = ft.TextField(
            label="Tìm kiếm theo tên, năm sinh, đơn vị, số hiệu...",
            prefix_icon=ft.Icons.SEARCH,
            on_change=self.filter_allowance_list,
            border_radius=12,
            bgcolor=ft.Colors.WHITE,
            border_color=ft.Colors.BLUE_200,
            focused_border_color=ft.Colors.BLUE_600,
            dense=True,
            expand=True
        )
        self.search_button = ft.ElevatedButton(
            "Tìm kiếm",
            icon=ft.Icons.SEARCH,
            on_click=self.filter_allowance_list,
            style=ft.ButtonStyle(
                bgcolor=ft.Colors.BLUE_800, 
                color=ft.Colors.WHITE, 
                shape=ft.RoundedRectangleBorder(radius=12),
                elevation=2
            )
        )
        self.list_container = ft.ListView(spacing=8, expand=True)

        # UI cho Tab 2: Thống Kê
        self.stats_search_field = ft.TextField(
            label="Tìm theo tên, số hiệu, đơn vị...",
            prefix_icon=ft.Icons.SEARCH,
            on_change=self.apply_statistics_filter,
            border_radius=10,
            bgcolor=ft.Colors.WHITE,
            border_color=ft.Colors.BLUE_200,
            focused_border_color=ft.Colors.BLUE_600,
            dense=True,
            expand=True
        )
        self.filter_unit_dd = ft.Dropdown(label="Đơn vị hiện tại", width=180, on_change=self.apply_statistics_filter, dense=True, border_radius=10, bgcolor=ft.Colors.WHITE)
        self.filter_rate_dd = ft.Dropdown(label="Mức hưởng", width=120, on_change=self.apply_statistics_filter, dense=True, border_radius=10, bgcolor=ft.Colors.WHITE)
        self.filter_status_dd = ft.Dropdown(
            label="Trạng thái hạn",
            width=160,
            options=[
                ft.dropdown.Option("Tất cả"),
                ft.dropdown.Option("Đến hạn/Quá hạn"),
                ft.dropdown.Option("Còn 1 tháng"),
                ft.dropdown.Option("Còn dưới 3 tháng"),
                ft.dropdown.Option("An toàn"),
            ],
            on_change=self.apply_statistics_filter,
            dense=True,
            border_radius=10,
            bgcolor=ft.Colors.WHITE
        )
        self.export_docx_btn = ft.ElevatedButton(
            "Xuất biểu mẫu Word",
            icon=ft.Icons.DESCRIPTION,
            on_click=self.trigger_docx_save_picker,
            style=ft.ButtonStyle(
                bgcolor=ft.Colors.BLUE_800,
                color=ft.Colors.WHITE,
                shape=ft.RoundedRectangleBorder(radius=10),
                elevation=2
            )
        )
        self.stats_total_text = ft.Text("Tổng số nhân sự: 0", size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_GREY_800)
        self.stats_container = ft.ListView(spacing=8, expand=True)

        # UI cho Tab 4: Điều Chỉnh
        self.adjust_search_field = ft.TextField(
            label="Lọc nhanh nhân sự cần chỉnh sửa dữ liệu...",
            prefix_icon=ft.Icons.EDIT,
            on_change=self.filter_adjustment_list,
            border_radius=12,
            bgcolor=ft.Colors.WHITE,
            border_color=ft.Colors.TEAL_200,
            focused_border_color=ft.Colors.TEAL_600,
            dense=True,
            expand=True
        )
        self.adjust_table_container = ft.Column(scroll=ft.ScrollMode.ALWAYS, expand=True)

        self.create_ui()

    def create_ui(self):
        header = ft.Container(
            content=ft.Row([
                ft.Row([
                    ft.Icon(ft.Icons.LAYERS, color=ft.Colors.WHITE, size=28),
                    ft.Text("HỆ THỐNG QUẢN LÝ THU HÚT", size=18, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                ], spacing=10),
                ft.ElevatedButton(
                    "Load Excel",
                    icon=ft.Icons.UPLOAD_FILE,
                    on_click=self.trigger_file_picker,
                    style=ft.ButtonStyle(
                        bgcolor=ft.Colors.AMBER_600, 
                        color=ft.Colors.WHITE, 
                        shape=ft.RoundedRectangleBorder(radius=10),
                        elevation=3
                    )
                )
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            gradient=ft.LinearGradient(
                begin=ft.alignment.top_left,
                end=ft.alignment.bottom_right,
                colors=[ft.Colors.BLUE_900, ft.Colors.INDIGO_700]
            ),
            padding=ft.padding.symmetric(horizontal=20, vertical=15),
            border_radius=12,
            shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.BLUE_GREY_200, offset=ft.Offset(0, 3))
        )

        # Tab 1 Layout
        tab_danh_sach = ft.Container(
            content=ft.Column([
                ft.Row([self.search_field, self.search_button], spacing=10),
                self.build_table_header(),
                ft.Container(content=self.list_container, expand=True)
            ], spacing=10),
            padding=10
        )

        # Tab 2 Layout
        tab_thong_ke = ft.Container(
            content=ft.Column([
                ft.Container(
                    content=ft.Row([
                        self.stats_search_field, 
                        self.filter_unit_dd, 
                        self.filter_rate_dd, 
                        self.filter_status_dd,
                        self.export_docx_btn
                    ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=10, bgcolor=ft.Colors.BLUE_GREY_100, border_radius=12
                ),
                ft.Row([
                    ft.Icon(ft.Icons.ANALYTICS, color=ft.Colors.BLUE_700, size=20),
                    self.stats_total_text
                ], spacing=8),
                ft.Container(content=self.stats_container, expand=True)
            ], spacing=12),
            padding=10
        )

        # Tab 3 Layout
        tab_huong_dan = ft.Container(
            content=ft.Column([
                ft.Text("VĂN BẢN HƯỚNG DẪN", size=15, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_GREY_900),
                ft.Divider(height=1, color=ft.Colors.GREY_300),
                # ft.Row([
                #     self.build_guidance_card("Hướng dẫn 01/HD-BCA", "Thực hiện chế độ phụ cấp thu hút và điều động cán bộ địa bàn trọng điểm.", "Bộ Công An", ft.Colors.RED_ACCENT_700),
                #     self.build_guidance_card("Nghị định 76/2019/NĐ-CP", "Chính sách đối với cán bộ, công chức, viên chức công tác ở vùng đặc biệt khó khăn.", "Chính Phủ", ft.Colors.BLUE_ACCENT_700),
                # ], spacing=15),
                # ft.Row([
                #     self.build_guidance_card("Thông tư 14/TT-BNV", "Hướng dẫn vạch định ranh giới và thời hạn tính mốc hưởng thu hút vùng biên giới.", "Bộ Nội Vụ", ft.Colors.TEAL_700),
                #     self.build_guidance_card("Quyết định số 22/QĐ-HĐND", "Phê duyệt bổ sung kinh phí hỗ trợ phụ cấp thu hút cán bộ năm hành chính.", "Hội Đồng Nhân Dân", ft.Colors.ORANGE_ACCENT_700),
                # ], spacing=15),
            ], spacing=15, scroll=ft.ScrollMode.ALWAYS),
            padding=15
        )

        # Tab 4 Layout
        tab_dieu_chinh = ft.Container(
            content=ft.Column([
                ft.Row([self.adjust_search_field]),
                ft.Text("💡 Mẹo: Nhấp vào bất kỳ ô nào để trực tiếp sửa đổi.", size=12, italic=True, color=ft.Colors.TEAL_800),
                ft.Container(
                    content=self.adjust_table_container,
                    expand=True,
                    border=ft.border.all(1, ft.Colors.GREY_300),
                    border_radius=8,
                    bgcolor=ft.Colors.WHITE
                )
            ], spacing=10),
            padding=10
        )

        self.tabs_control = ft.Tabs(
            selected_index=0,
            animation_duration=200,
            tabs=[
                ft.Tab(text="1. Danh sách tổng hợp", icon=ft.Icons.LIST_ALT, content=tab_danh_sach),
                ft.Tab(text="2. Bộ lọc & Thống kê", icon=ft.Icons.BAR_CHART, content=tab_thong_ke),
                ft.Tab(text="3. Văn bản hướng dẫn", icon=ft.Icons.BOOKMARK, content=tab_huong_dan),
                ft.Tab(text="4. Điều chỉnh", icon=ft.Icons.GRID_ON, content=tab_dieu_chinh),
            ],
            expand=True
        )

        self.page.add(ft.Column([header, ft.Container(height=5), self.tabs_control], expand=True, spacing=0))

    def build_guidance_card(self, title, desc, author, accent_color):
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Icon(ft.Icons.ARTICLE, color=accent_color, size=20),
                    ft.Text(title, size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_GREY_900)
                ], spacing=8),
                ft.Text(desc, size=13, color=ft.Colors.BLUE_GREY_700, max_lines=2, overflow=ft.TextOverflow.ELLIPSIS),
                ft.Divider(height=8, color=ft.Colors.GREY_100),
                ft.Row([
                    ft.Container(
                        content=ft.Text(author, size=10, color=ft.Colors.WHITE, weight=ft.FontWeight.BOLD),
                        bgcolor=accent_color, padding=ft.padding.symmetric(horizontal=8, vertical=3), border_radius=5
                    ),
                    ft.TextButton("Xem chi tiết", icon=ft.Icons.OPEN_IN_NEW, style=ft.ButtonStyle(color=accent_color))
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
            ], spacing=8),
            bgcolor=ft.Colors.WHITE,
            padding=14,
            border_radius=10,
            border=ft.border.all(1, ft.Colors.GREY_200),
            expand=True,
            shadow=ft.BoxShadow(blur_radius=4, color=ft.Colors.GREY_300, offset=ft.Offset(0, 2))
        )

    def build_table_header(self):
        return ft.Container(
            content=ft.Row([
                ft.Text("Họ và tên", size=12, weight=ft.FontWeight.BOLD, width=180, color=ft.Colors.WHITE),
                ft.Text("Số hiệu", size=12, weight=ft.FontWeight.BOLD, width=100, color=ft.Colors.WHITE),
                ft.Text("Đơn vị hiện tại", size=12, weight=ft.FontWeight.BOLD, width=160, color=ft.Colors.WHITE),
                ft.Text("Mức hiện tại", size=12, weight=ft.FontWeight.BOLD, width=100, color=ft.Colors.WHITE),
                ft.Text("Các mốc thời gian hưởng", size=12, weight=ft.FontWeight.BOLD, width=200, color=ft.Colors.WHITE),
                ft.Text("Trạng thái hạn", size=12, weight=ft.FontWeight.BOLD, expand=True, text_align=ft.TextAlign.RIGHT, color=ft.Colors.WHITE),
            ]),
            gradient=ft.LinearGradient(colors=[ft.Colors.BLUE_GREY_800, ft.Colors.BLUE_GREY_700]),
            padding=ft.padding.symmetric(horizontal=12, vertical=12),
            border_radius=8
        )

    def trigger_file_picker(self, e):
        file_picker = ft.FilePicker(on_result=self.process_excel_file)
        self.page.overlay.append(file_picker)
        self.page.update()
        file_picker.pick_files(allowed_extensions=["xlsx"])

    def calculate_months_left(self, start_str: str, end_str: str) -> tuple[str, str, str]:
        if not end_str or end_str == "Chưa cập nhật":
            return "GREEN", "Chưa cập nhật hạn cuối", ft.Colors.GREEN_700

        try:
            end_date = datetime.strptime(end_str, '%d/%m/%Y')
            today = datetime.now()

            if end_date <= today: return "BLACK", "Đến hạn/Quá hạn", ft.Colors.BLACK

            diff = relativedelta(end_date, today)
            total_months = diff.years * 12 + diff.months

            if total_months <= 1: return "RED", "Còn 1 tháng", ft.Colors.RED_600
            elif total_months <= 3: return "ORANGE", f"Còn {total_months} tháng", ft.Colors.ORANGE_700
            else: return "GREEN", f"Còn {total_months} tháng", ft.Colors.GREEN_700
        except:
            return "GREY", "Chưa xác định", ft.Colors.GREY_600

    def process_excel_file(self, e: ft.FilePickerResultEvent):
        if not e.files: return
        try:
            path = e.files[0].path
            sheets_dict = pd.read_excel(path, sheet_name=None)
            self.master_data.clear()
            self.extra_headers.clear()

            for sheet_name, df in sheets_dict.items():
                if df.empty: continue
                df.columns = [str(c).strip() for c in df.columns]

                name_col = find_column_by_keywords(df.columns, ['ho va ten', 'ho ten', 'hoten', 'ten'])
                yob_col = find_column_by_keywords(df.columns, ['nam sinh', 'namsinh', 'ngay sinh', 'ngaysinh'])
                cand_col = find_column_by_keywords(df.columns, ['so hieu cand', 'so hieu', 'sh cand', 'so hieu quan nhan'])
                unit_col = find_column_by_keywords(df.columns, ['don vi hien tai', 'don vi', 'phong ban', 'donvi'])
                rate_col = find_column_by_keywords(df.columns, ['muc huong', 'ty le', 'he so', 'muc phu cap'])
                start_date_col = find_column_by_keywords(df.columns, ['bat dau huong', 'thoi gian nhan cong tactu ngay', 'tu ngay'])
                end_date_col = find_column_by_keywords(df.columns, ['ket thuc huong', 'han huong', 'den ngay'])
                total_months_col = find_column_by_keywords(df.columns, ['tong so thang', 'tong thang', 'so thang', 'thoi gian'])

                if not name_col: continue
                last_name, last_yob, last_cand, last_unit = "", "Chưa cập nhật", "Chưa cập nhật", "Chưa cập nhật"

                yob_regex = re.compile(r"\((19\d{2}|20\d{2})\)")
                yob_replace_regex = re.compile(r"\s*\(19\d{2}\)|\s*\(20\d{2}\)")
                today_date = datetime.now() 

                for row_tuple in df.itertuples(index=False):
                    row = dict(zip(df.columns, row_tuple))

                    raw_name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
                    
                    extracted_yob = ""
                    if raw_name and raw_name.lower() != 'nan':
                        # DÙNG REGEX ĐÃ BIÊN DỊCH TRƯỚC ĐỂ TĂNG TỐC
                        match_yob = yob_regex.search(raw_name) 
                        if match_yob:
                            extracted_yob = match_yob.group(1)
                            raw_name = yob_replace_regex.sub("", raw_name).strip()
                        if match_yob:
                            extracted_yob = match_yob.group(1)
                            raw_name = re.sub(r"\s*\(19\d{2}\)|\s*\(20\d{2}\)", "", raw_name).strip()

                    raw_yob = ""
                    if yob_col and pd.notna(row[yob_col]):
                        val_yob = str(row[yob_col]).strip()
                        if val_yob.endswith('.0'): val_yob = val_yob[:-2]
                        raw_yob = val_yob
                    
                    if not raw_yob or raw_yob.lower() in ['nan', '', 'nat', '-', '.', '0']:
                        raw_yob = extracted_yob if extracted_yob else "Chưa cập nhật"

                    if raw_name and raw_name.lower() != 'nan':
                        last_name = raw_name
                        last_yob = raw_yob if (raw_yob and raw_yob.lower() != 'nan') else "Chưa cập nhật"
                        last_cand = "Chưa cập nhật"
                        last_unit = "Chưa cập nhật"
                    else:
                        raw_name = last_name
                        raw_yob = last_yob

                    if not raw_name or raw_name.lower() in ['nan', '']: continue

                    norm_key = f"{normalize_text(raw_name)}_{normalize_text(raw_yob)}"

                    raw_cand = str(row[cand_col]).strip() if cand_col and pd.notna(row[cand_col]) else ""
                    if raw_cand and raw_cand.lower() != 'nan': last_cand = raw_cand
                    else: raw_cand = last_cand

                    raw_unit = str(row[unit_col]).strip() if unit_col and pd.notna(row[unit_col]) else ""
                    if raw_unit and raw_unit.lower() != 'nan': last_unit = clean_unit_name(raw_unit)
                    else: raw_unit = last_unit

                    rate_val = str(row[rate_col]).strip() if rate_col and pd.notna(row[rate_col]) else ""
                    if not rate_val or rate_val.lower() in ['nan', '', '-', '.', '0']: rate_val = "70%"

                    start_val = clean_cell_to_date_str(row[start_date_col]) if start_date_col else ""
                    end_val = clean_cell_to_date_str(row[end_date_col]) if end_date_col else ""

                    if start_val and not end_val:
                        try:
                            start_dt = datetime.strptime(start_val, '%d/%m/%Y')
                            end_dt = start_dt + relativedelta(years=5)
                            end_val = end_dt.strftime('%d/%m/%Y')
                        except: pass

                    smart_months_val = "Chưa cập nhật"
                    if start_val and start_val != "Chưa cập nhật":
                        try:
                            start_dt = datetime.strptime(start_val, '%d/%m/%Y')
                            today = today_date
                            end_point = today
                            if end_val and end_val != "Chưa cập nhật":
                                end_dt = datetime.strptime(end_val, '%d/%m/%Y')
                                end_point = min(today, end_dt)

                            if end_point >= start_dt:
                                diff = relativedelta(end_point, start_dt)
                                total_m = diff.years * 12 + diff.months
                                breakdown = []
                                if diff.years > 0: breakdown.append(f"{diff.years} năm")
                                if diff.months > 0: breakdown.append(f"{diff.months} tháng")
                                smart_months_val = f"{total_m} tháng ({' '.join(breakdown)})" if breakdown else f"{total_m} tháng"
                        except: pass

                    row_details = {}
                    unnamed_idx = 1
                    chuthich_lines = []  # Danh sách tạm để gom các dòng chú thích

                    for k, v in row.items():
                        k_str = str(k).strip()
                        # Kiểm tra xem đây có phải cột không có header (Unnamed) không
                        is_unnamed = "unnamed" in k_str.lower() or not k_str
                        
                        v_str = str(v).strip()
                        # Kiểm tra ô có dữ liệu thực tế hay không
                        has_value = not (pd.isna(v) or v_str.lower() in ['nan', 'nat', '-', '.', '0', ''])

                        # NẾU LÀ CỘT UNNAMED: Gom vào cột "Chú thích" tổng hợp (nếu có dữ liệu)
                        if is_unnamed:
                            if has_value:
                                chuthich_lines.append(f"- Ghi chú {unnamed_idx}: {v_str}")
                            unnamed_idx += 1
                            continue  # Bỏ qua các xử lý core_cols hoặc extra_headers ở dưới cho cột này

                        # NẾU LÀ CỘT CÓ TÊN: Xử lý phân loại bình thường
                        if name_col and k == name_col: row_details[k_str] = raw_name
                        elif yob_col and k == yob_col: row_details[k_str] = raw_yob
                        elif cand_col and k == cand_col: row_details[k_str] = raw_cand
                        elif unit_col and k == unit_col: row_details[k_str] = raw_unit
                        elif rate_col == k: row_details[k_str] = rate_val
                        elif start_date_col and k == start_date_col: row_details[k_str] = start_val if start_val else "Chưa cập nhật"
                        elif end_date_col and k == end_date_col: row_details[k_str] = end_val if end_val else "Chưa cập nhật"
                        elif total_months_col and k == total_months_col: row_details[k_str] = smart_months_val
                        else:
                            row_details[k_str] = v_str if has_value else "Chưa cập nhật"

                        # Thu thập tự động các cột phụ CÓ TÊN (không nằm trong nhóm dữ liệu chính)
                        core_cols = [
                            str(name_col).strip() if name_col else "",
                            str(yob_col).strip() if yob_col else "",
                            str(cand_col).strip() if cand_col else "",
                            str(unit_col).strip() if unit_col else "",
                            str(rate_col).strip() if rate_col else "",
                            str(start_date_col).strip() if start_date_col else "",
                            str(end_date_col).strip() if end_date_col else ""
                        ]
                        if k_str not in core_cols and k_str not in self.extra_headers:
                            self.extra_headers.append(k_str)

                    # SAU KHI DUYỆT HẾT CÁC CỘT: Tổng hợp danh sách chú thích vào row_details
                    if chuthich_lines:
                        row_details["Chú thích"] = "\n".join(chuthich_lines)
                    else:
                        row_details["Chú thích"] = "Không có"  # Hoặc để trống "" tùy bạn

                    # Đưa duy nhất một cột "Chú thích" vào danh sách hiển thị
                    if "Chú thích" not in self.extra_headers:
                        self.extra_headers.append("Chú thích")

                    if not total_months_col:
                        row_details["Tổng số tháng"] = smart_months_val
                        if "Tổng số tháng" not in self.extra_headers:
                            self.extra_headers.append("Tổng số tháng")

                    if start_val and end_val: duration_text = f"{start_val} → {end_val}"
                    elif start_val: duration_text = f"Từ {start_val}"
                    elif end_val: duration_text = f"Đến {end_val}"
                    else: duration_text = "Chưa cập nhật"

                    color_code, badge_desc, color_flet = self.calculate_months_left(start_val, end_val)

                    if norm_key not in self.master_data:
                        self.master_data[norm_key] = {
                            "display_name": raw_name,
                            "yob": raw_yob,
                            "cand_id": raw_cand if raw_cand else "Chưa cập nhật",
                            "unit": raw_unit if raw_unit else "Chưa cập nhật",
                            "all_records": []
                        }

                    is_duplicate = False
                    for existing in self.master_data[norm_key]["all_records"]:
                        if (existing["start_date"] == start_val and 
                            existing["end_date"] == end_val and 
                            existing["rate"] == rate_val and 
                            existing["unit"] == (raw_unit if raw_unit else "Chưa cập nhật")):
                            is_duplicate = True
                            break
                    
                    if not is_duplicate:
                        self.master_data[norm_key]["all_records"].append({
                            "sheet": sheet_name,
                            "start_date": start_val,
                            "end_date": end_val,
                            "rate": rate_val,
                            "duration_str": duration_text,
                            "full_row": row_details,
                            "unit": raw_unit if raw_unit else "Chưa cập nhật",
                            "cand_id": raw_cand if raw_cand else "Chưa cập nhật",
                            "status_color_code": color_code,
                            "status_badge_desc": badge_desc,
                            "status_color_flet": color_flet
                        })

            name_to_keys = {}
            for k, emp in self.master_data.items():
                norm_name = normalize_text(emp["display_name"])
                if norm_name not in name_to_keys:
                    name_to_keys[norm_name] = []
                name_to_keys[norm_name].append(k)

            for norm_name, keys in name_to_keys.items():
                if len(keys) > 1:
                    with_yob, without_yob = [], []
                    for k in keys:
                        yob_clean = str(self.master_data[k]["yob"]).strip().lower()
                        if yob_clean in ["chưa cập nhật", "nan", ""]: without_yob.append(k)
                        else: with_yob.append(k)

                    if len(with_yob) == 1 and len(without_yob) > 0:
                        target_key = with_yob[0]
                        for w_k in without_yob:
                            self.master_data[target_key]["all_records"].extend(self.master_data[w_k]["all_records"])
                            del self.master_data[w_k]
                    elif len(with_yob) == 0 and len(without_yob) > 1:
                        target_key = without_yob[0]
                        for w_k in without_yob[1:]:
                            self.master_data[target_key]["all_records"].extend(self.master_data[w_k]["all_records"])
                            del self.master_data[w_k]

            def get_sort_date(rec):
                s_date = rec.get("start_date", "")
                if not s_date or s_date == "Chưa cập nhật": return datetime.min
                try: return datetime.strptime(s_date, "%d/%m/%Y")
                except: return datetime.min

            for norm_key, emp in self.master_data.items():
                emp["all_records"].sort(key=get_sort_date)
                final_unit, final_cand = "Chưa cập nhật", "Chưa cập nhật"
                for rec in reversed(emp["all_records"]):
                    u = rec.get("unit", "Chưa cập nhật")
                    if u and u != "Chưa cập nhật" and final_unit == "Chưa cập nhật": final_unit = u
                    c = rec.get("cand_id", "Chưa cập nhật")
                    if c and c != "Chưa cập nhật" and final_cand == "Chưa cập nhật": final_cand = c
                emp["unit"] = final_unit
                emp["cand_id"] = final_cand

            unique_units = set(emp["unit"] for emp in self.master_data.values() if emp["unit"] != "Chưa cập nhật")
            unique_rates = set(rec["rate"] for emp in self.master_data.values() for rec in emp["all_records"] if rec["rate"] != "Chưa cập nhật")

            self.filter_unit_dd.options = [ft.dropdown.Option("Tất cả")] + [
                ft.dropdown.Option(u) for u in sorted(unique_units, key=lambda x: str(x).lower())
            ]
            self.filter_rate_dd.options = [ft.dropdown.Option("Tất cả")] + [
                ft.dropdown.Option(r) for r in sorted(unique_rates, key=lambda x: str(x).lower())
            ]
            self.filter_unit_dd.value = "Tất cả"
            self.filter_rate_dd.value = "Tất cả"
            self.filter_status_dd.value = "Tất cả"

            self.filtered_data = self.master_data.copy()
            self.render_allowance_list(append=False)
            self.render_statistics_list(append=False)
            self.render_adjustment_list(append=False)

            self.page.open(ft.SnackBar(ft.Text("Đồng bộ dữ liệu thành công!"), bgcolor=ft.Colors.GREEN_600))
        except Exception as ex:
            self.page.open(ft.SnackBar(ft.Text(f"Lỗi đọc file: {str(ex)}"), bgcolor=ft.Colors.RED_600))
        self.page.update()

    def get_clean_display_title(self, emp) -> str:
        name = emp.get("display_name", "")
        yob = str(emp.get("yob", "")).strip()
        if yob and yob.lower() not in ["chưa cập nhật", "nan", ""]: 
            return f"{name} ({yob})"
        return name

    def show_history_popup(self, person_key):
        emp = self.master_data.get(person_key)
        if not emp or not emp["all_records"]: return

        display_yob = emp['yob'] if emp['yob'] != "Chưa cập nhật" else ""
        info_profile = ft.Container(
            content=ft.ResponsiveRow([
                ft.Column([
                    ft.Text("Họ và tên", size=11, color=ft.Colors.BLUE_GREY_400, weight=ft.FontWeight.W_500),
                    ft.Text(emp['display_name'], size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900)
                ], col={"xs": 12, "sm": 6, "md": 3}),
                ft.Column([
                    ft.Text("Năm sinh", size=11, color=ft.Colors.BLUE_GREY_400, weight=ft.FontWeight.W_500),
                    ft.Text(display_yob, size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900)
                ], col={"xs": 12, "sm": 6, "md": 2}),
                ft.Column([
                    ft.Text("Số hiệu CAND", size=11, color=ft.Colors.BLUE_GREY_400, weight=ft.FontWeight.W_500),
                    ft.Text(emp['cand_id'], size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900)
                ], col={"xs": 12, "sm": 6, "md": 3}),
                ft.Column([
                    ft.Text("Đơn vị hiện tại", size=11, color=ft.Colors.BLUE_GREY_400, weight=ft.FontWeight.W_500),
                    ft.Text(emp['unit'], size=14, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900, overflow=ft.TextOverflow.ELLIPSIS)
                ], col={"xs": 12, "sm": 6, "md": 4}),
            ], spacing=10),
            bgcolor=ft.Colors.BLUE_GREY_50, padding=15, border_radius=8, margin=ft.margin.only(bottom=15)
        )

        sample_row = emp["all_records"][0]["full_row"]
        all_headers = list(sample_row.keys())
        exclude_kws = ['ho va ten', 'ho ten', 'hoten', 'ten', 'nam sinh', 'namsinh', 'ngay sinh', 'ngaysinh', 'so hieu cand', 'so hieu', 'sh cand', 'so hieu quan nhan']
        filtered_headers = [h for h in all_headers if normalize_text(h) not in exclude_kws]

        extended_headers = filtered_headers + ["Trạng thái hạn"]
        data_columns = [ft.DataColumn(ft.Text(h, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_GREY_700, size=12)) for h in extended_headers]

        data_rows = []
        for rec in emp["all_records"]:
            badge_desc = rec.get("status_badge_desc", "Chưa rõ")
            color_flet = rec.get("status_color_flet", ft.Colors.GREY_600)

            row_cells = []
            for h in filtered_headers:
                v = rec["full_row"].get(h, "Chưa cập nhật")
                row_cells.append(ft.DataCell(ft.Text(str(v), size=12, color=ft.Colors.BLUE_GREY_900)))

            row_cells.append(ft.DataCell(
                ft.Container(
                    content=ft.Text(badge_desc, color=ft.Colors.WHITE, size=10, weight=ft.FontWeight.BOLD),
                    bgcolor=color_flet, padding=ft.padding.symmetric(horizontal=8, vertical=3), border_radius=4
                )
            ))
            data_rows.append(ft.DataRow(cells=row_cells))

        detail_table = ft.DataTable(
            columns=data_columns, rows=data_rows, heading_row_color=ft.Colors.GREY_100, divider_thickness=1,
            horizontal_lines=ft.BorderSide(0.5, ft.Colors.GREY_300), vertical_lines=ft.BorderSide(0.5, ft.Colors.GREY_200),
        )

        scrollable_container = ft.Column(expand=True, scroll=ft.ScrollMode.ALWAYS, controls=[ft.Row(scroll=ft.ScrollMode.ALWAYS, controls=[detail_table])])

        popup_layout = ft.Column([
            info_profile,
            ft.Text("📅 Lịch sử điều chỉnh:", size=12, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_GREY_600),
            ft.Container(content=scrollable_container, expand=True)
        ], spacing=5, expand=True)

        def redirect_to_excel_mode(e):
            self.page.close(dialog)
            self.tabs_control.selected_index = 3 
            self.adjust_search_field.value = emp['display_name']
            self.render_adjustment_list(append=False)
            self.page.update()

        dialog = ft.AlertDialog(
            title=ft.Row([
                ft.Icon(ft.Icons.HISTORY, color=ft.Colors.BLUE_800, size=24),
                ft.Text("Chi tiết", size=16, weight=ft.FontWeight.BOLD)
            ], spacing=8),
            content=ft.Container(content=popup_layout, width=1050, height=520),
            actions=[
                ft.ElevatedButton("Điều chỉnh/Cập nhật thông tin", on_click=redirect_to_excel_mode, style=ft.ButtonStyle(bgcolor=ft.Colors.TEAL_700, color=ft.Colors.WHITE, shape=ft.RoundedRectangleBorder(radius=8))),
                ft.TextButton("Đóng ", on_click=lambda e: self.page.close(dialog), style=ft.ButtonStyle(color=ft.Colors.BLUE_GREY_700))
            ]
        )
        self.page.open(dialog)

    def load_more_data(self, e):
        self.page_number += 1
        self.render_allowance_list(append=True)

    def render_allowance_list(self, append=False):
        if not append:
            self.list_container.controls.clear()
            self.page_number = 1

        if not self.filtered_data:
            self.list_container.controls.append(ft.Text("Không tìm thấy kết quả phù hợp.", color=ft.Colors.GREY_500, italic=True))
            self.page_number = 1
            self.page.update()
            return

        if append and len(self.list_container.controls) > 0:
            if getattr(self.list_container.controls[-1], "key", None) == "load_more_btn":
                self.list_container.controls.pop()

        items = list(self.filtered_data.items())
        start_idx = (self.page_number - 1) * self.per_page
        end_idx = start_idx + self.per_page
        page_items = items[start_idx:end_idx]

        for norm_key, emp in page_items:
            duration_views = [ft.Text(rec["duration_str"], size=13) for rec in emp["all_records"]]
            latest_rec = emp["all_records"][-1]
            badge_desc = latest_rec.get("status_badge_desc", "Chưa rõ")
            color_flet = latest_rec.get("status_color_flet", ft.Colors.GREY_600)
            display_title = self.get_clean_display_title(emp)

            row_card = ft.Container(
                content=ft.Row([
                    ft.Container(content=ft.Text(display_title, size=13, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_800, overflow=ft.TextOverflow.ELLIPSIS), width=180),
                    ft.Text(emp["cand_id"], size=13, width=100),
                    ft.Text(emp["unit"], size=13, width=160, overflow=ft.TextOverflow.ELLIPSIS),
                    ft.Container(content=ft.Text(latest_rec["rate"], size=13), width=100),
                    ft.Column(controls=duration_views, width=200, spacing=6),
                    ft.Container(
                        content=ft.Container(
                            content=ft.Text(badge_desc, color=ft.Colors.WHITE, size=11, weight=ft.FontWeight.W_500),
                            bgcolor=color_flet, padding=ft.padding.symmetric(horizontal=10, vertical=2), border_radius=4,
                        ),
                        expand=True, alignment=ft.alignment.center_right
                    )
                ], vertical_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=ft.Colors.WHITE, padding=ft.padding.symmetric(horizontal=12, vertical=12), border_radius=8,
                border=ft.border.all(0.5, ft.Colors.GREY_200), on_click=lambda e, k=norm_key: self.show_history_popup(k),
                shadow=ft.BoxShadow(blur_radius=4, color=ft.Colors.GREY_200, offset=ft.Offset(0, 1))
            )
            self.list_container.controls.append(row_card)

        if len(items) > end_idx:
            load_more_btn = ft.Container(
                key="load_more_btn",
                content=ft.ElevatedButton(
                    f"Xem thêm ({len(items) - end_idx} nhân sự trong kết quả tìm kiếm)...",
                    icon=ft.Icons.ARROW_DROP_DOWN_CIRCLE_ROUNDED,
                    on_click=self.load_more_data,
                    style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_50, color=ft.Colors.BLUE_800)
                ),
                alignment=ft.alignment.center, padding=ft.padding.only(top=10, bottom=10)
            )
            self.list_container.controls.append(load_more_btn)
        self.page.update()

    def filter_allowance_list(self, e):
        q = normalize_text(self.search_field.value)
        if not q: self.filtered_data = self.master_data.copy()
        else:
            self.filtered_data = {}
            for k, emp in self.master_data.items():
                if (q in normalize_text(emp["display_name"]) or q in normalize_text(emp["cand_id"]) or q in normalize_text(emp["unit"]) or q in normalize_text(emp["yob"])):
                    self.filtered_data[k] = emp
        self.render_allowance_list(append=False)

    def load_more_stats_data(self, e):
        self.stats_page_number += 1
        self.render_statistics_list(append=True)

    def render_statistics_list(self, append=False):
        if not append:
            self.stats_container.controls.clear()
            self.stats_page_number = 1
            self.stats_filtered_items = []

            q_stats = normalize_text(self.stats_search_field.value)
            unit_f = self.filter_unit_dd.value
            rate_f = self.filter_rate_dd.value
            status_f = self.filter_status_dd.value

            for norm_key, emp in self.master_data.items():
                if q_stats:
                    if not (q_stats in normalize_text(emp["display_name"]) or 
                            q_stats in normalize_text(emp["cand_id"]) or 
                            q_stats in normalize_text(emp["unit"]) or 
                            q_stats in normalize_text(emp["yob"])):
                        continue

                if unit_f and unit_f != "Tất cả" and emp["unit"] != unit_f: continue
                
                match_found = True
                if (rate_f and rate_f != "Tất cả") or (status_f and status_f != "Tất cả"):
                    has_rate, has_status = False, False
                    for rec in emp["all_records"]:
                        color_code = rec.get("status_color_code", "GREY")
                        if rate_f and rate_f != "Tất cả" and rec["rate"] == rate_f: has_rate = True
                        if status_f and status_f != "Tất cả":
                            if status_f == "Đến hạn/Quá hạn" and color_code == "BLACK": has_status = True
                            if status_f == "Còn 1 tháng" and color_code == "RED": has_status = True
                            if status_f == "Còn dưới 3 tháng" and color_code == "ORANGE": has_status = True
                            if status_f == "An toàn" and color_code == "GREEN": has_status = True

                    if rate_f and rate_f != "Tất cả" and not has_rate: match_found = False
                    if status_f and status_f != "Tất cả" and not has_status: match_found = False

                if not match_found: continue
                self.stats_filtered_items.append((norm_key, emp))
            self.stats_total_text.value = f"📊 Tổng số nhân sự thỏa mãn điều kiện: {len(self.stats_filtered_items)} người"

        if append and len(self.stats_container.controls) > 0:
            if getattr(self.stats_container.controls[-1], "key", None) == "load_more_stats_btn":
                self.stats_container.controls.pop()

        start_idx = (self.stats_page_number - 1) * self.per_page
        end_idx = start_idx + self.per_page
        page_items = self.stats_filtered_items[start_idx:end_idx]

        for norm_key, emp in page_items:
            latest_rec = emp["all_records"][-1]
            badge_desc = latest_rec.get("status_badge_desc", "Chưa rõ")
            color_flet = latest_rec.get("status_color_flet", ft.Colors.GREY_600)
            display_title = self.get_clean_display_title(emp)

            stats_row = ft.Container(
                content=ft.Row([
                    ft.Text(display_title, size=13, weight=ft.FontWeight.W_500, width=180),
                    ft.Text(emp["unit"], size=13, width=200),
                    ft.Container(content=ft.Text(latest_rec["rate"], size=13), width=100),
                    ft.Container(
                        content=ft.Container(
                            content=ft.Text(badge_desc, color=ft.Colors.WHITE, size=10),
                            bgcolor=color_flet, padding=ft.padding.symmetric(horizontal=8, vertical=2), border_radius=4
                        ),
                        expand=True, alignment=ft.alignment.center_right
                    )
                ], vertical_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=ft.Colors.WHITE, padding=12, border_radius=6,
                on_click=lambda e, k=norm_key: self.show_history_popup(k),
                border=ft.border.all(0.5, ft.Colors.GREY_200)
            )
            self.stats_container.controls.append(stats_row)

        if len(self.stats_filtered_items) > end_idx:
            load_more_stats_btn = ft.Container(
                key="load_more_stats_btn",
                content=ft.ElevatedButton(
                    f"Xem thêm ({len(self.stats_filtered_items) - end_idx} nhân sự)...",
                    icon=ft.Icons.ARROW_DROP_DOWN_CIRCLE_ROUNDED,
                    on_click=self.load_more_stats_data,
                    style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_50, color=ft.Colors.BLUE_800)
                ),
                alignment=ft.alignment.center, padding=ft.padding.only(top=10, bottom=10)
            )
            self.stats_container.controls.append(load_more_stats_btn)
        self.page.update()

    def apply_statistics_filter(self, e):
        self.render_statistics_list(append=False)

    def trigger_docx_save_picker(self, e):
        if not self.stats_filtered_items:
            self.page.open(ft.SnackBar(ft.Text("Không có dữ liệu thỏa mãn bộ lọc hiện tại để xuất Word!"), bgcolor=ft.Colors.ORANGE_600))
            return
        
        docx_picker = ft.FilePicker(on_result=self.generate_docx_callback)
        self.page.overlay.append(docx_picker)
        self.page.update()
        docx_picker.save_file(
            allowed_extensions=["docx"], 
            file_name=f"Bao_cao_phu_cap_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        )

    def generate_docx_callback(self, e: ft.FilePickerResultEvent):
        if not e.path: return
        try:
            from docx import Document
            from docx.shared import Pt, Inches
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            from docx.enum.table import WD_ALIGN_VERTICAL

            doc = Document()
            
            # Cấu hình Layout khổ giấy A4, lề chuẩn văn bản hành chính (0.6 inch ~ 1.5 cm)
            sections = doc.sections
            for section in sections:
                section.top_margin = Inches(0.6)
                section.bottom_margin = Inches(0.6)
                section.left_margin = Inches(0.6)
                section.right_margin = Inches(0.6)

            num_people = len(self.stats_filtered_items)

            # ================== TRƯỜNG HỢP 1: TRONG THỐNG KÊ CHỈ CÓ 01 NGƯỜI ==================
            if num_people == 1:
                _, emp = self.stats_filtered_items[0]
                
                # Khung Cơ quan chủ quản và Quốc hiệu Tiêu ngữ (Bảng ẩn viền)
                header_table = doc.add_table(rows=1, cols=2)
                header_table.autofit = False
                header_table.columns[0].width = Inches(3.0)
                header_table.columns[1].width = Inches(4.5)
                
                cell_left = header_table.cell(0, 0)
                p_left = cell_left.paragraphs[0]
                p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_l1 = p_left.add_run("CÔNG AN TỈNH CAO BẰNG\n")
                r_l1.font.name = 'Arial'
                r_l1.font.size = Pt(11)
                r_l2 = p_left.add_run("PHÒNG TỔ CHỨC CÁN BỘ")
                r_l2.font.name = 'Arial'
                r_l2.font.size = Pt(11)
                r_l2.font.bold = True
                r_l2.font.underline = True
                
                cell_right = header_table.cell(0, 1)
                p_right = cell_right.paragraphs[0]
                p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_r1 = p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
                r_r1.font.name = 'Arial'
                r_r1.font.size = Pt(11)
                r_r1.font.bold = True
                r_r2 = p_right.add_run("Độc lập – Tự do – Hạnh phúc")
                r_r2.font.name = 'Arial'
                r_r2.font.size = Pt(11)
                r_r2.font.bold = True
                r_r2.font.underline = True
                
                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(12)
                
                # Địa danh, ngày tháng năm
                p_date = doc.add_paragraph()
                p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                current_date_str = f"Cao Bằng, ngày {datetime.now().day} tháng {datetime.now().month} năm {datetime.now().year}"
                r_date = p_date.add_run(current_date_str)
                r_date.font.name = 'Arial'
                r_date.font.size = Pt(11)
                r_date.font.italic = True
                
                p_space2 = doc.add_paragraph()
                p_space2.paragraph_format.space_before = Pt(18)
                
                # Tiêu đề
                p_title = doc.add_paragraph()
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_title = p_title.add_run("QUÁ TRÌNH HƯỞNG PHỤ CẤP")
                r_title.font.name = 'Arial'
                r_title.font.size = Pt(14)
                r_title.font.bold = True
                
                p_space3 = doc.add_paragraph()
                p_space3.paragraph_format.space_before = Pt(18)
                
                # Thông tin định danh cá nhân (nằm ngoài bảng)
                sh_str = emp['cand_id'] if emp['cand_id'] != "Chưa cập nhật" else ""
                p_info1 = doc.add_paragraph()
                r_info1 = p_info1.add_run(f"Họ và tên: {emp['display_name']};        Số hiệu: {sh_str}")
                r_info1.font.name = 'Arial'
                r_info1.font.size = Pt(11)
                
                dv_str = emp['unit'] if emp['unit'] != "Chưa cập nhật" else ""
                p_info2 = doc.add_paragraph()
                r_info2 = p_info2.add_run(f"Đơn vị: {dv_str}")
                r_info2.font.name = 'Arial'
                r_info2.font.size = Pt(11)
                
                p_info3 = doc.add_paragraph()
                r_info3 = p_info3.add_run("Quá trình hưởng:")
                r_info3.font.name = 'Arial'
                r_info3.font.size = Pt(11)
                r_info3.font.bold = True
                
                # Bảng quá trình hưởng độc lập
                num_records = len(emp["all_records"])
                table = doc.add_table(rows=num_records + 2, cols=6, style='Table Grid')
                table.autofit = False
                
                col_widths = [Inches(0.5), Inches(2.2), Inches(1.1), Inches(1.1), Inches(1.5), Inches(1.1)]
                for row in table.rows:
                    for idx, width in enumerate(col_widths):
                        row.cells[idx].width = width
                
                headers = ["STT", "Đơn vị", "Thời gian bắt đầu", "Thời gian kết thúc", "Lý do kết thúc", "Thời gian"]
                for idx, h_text in enumerate(headers):
                    cell = table.cell(0, idx)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run(h_text)
                    r.font.name = 'Arial'
                    r.font.size = Pt(10)
                    r.font.bold = True
                
                total_months_sum = 0
                for idx, rec in enumerate(emp["all_records"]):
                    row_idx = idx + 1
                    stt = str(idx + 1)
                    unit = rec.get("unit", "")
                    if unit == "Chưa cập nhật": unit = ""
                    start = rec.get("start_date", "")
                    if start == "Chưa cập nhật": start = ""
                    end = rec.get("end_date", "")
                    if end == "Chưa cập nhật": end = ""
                    
                    reason = rec["full_row"].get("Lý do kết thúc", rec["full_row"].get("Lý do", ""))
                    if reason == "Chưa cập nhật": reason = ""
                        
                    duration = rec["full_row"].get("Tổng số tháng", rec["full_row"].get("Thời gian", ""))
                    if duration == "Chưa cập nhật": duration = ""
                    
                    match = re.search(r"(\d+)\s*tháng", str(duration).lower())
                    if match:
                        total_months_sum += int(match.group(1))
                    elif str(duration).isdigit():
                        total_months_sum += int(duration)
                        duration = f"{duration} tháng"
                    
                    vals = [stt, str(unit), str(start), str(end), str(reason), str(duration)]
                    alignments = [
                        WD_ALIGN_PARAGRAPH.CENTER,
                        WD_ALIGN_PARAGRAPH.LEFT,
                        WD_ALIGN_PARAGRAPH.CENTER,
                        WD_ALIGN_PARAGRAPH.CENTER,
                        WD_ALIGN_PARAGRAPH.LEFT,
                        WD_ALIGN_PARAGRAPH.CENTER
                    ]
                    
                    for col_idx, val in enumerate(vals):
                        cell = table.cell(row_idx, col_idx)
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        p = cell.paragraphs[0]
                        p.alignment = alignments[col_idx]
                        r = p.add_run(val)
                        r.font.name = 'Arial'
                        r.font.size = Pt(10)
                
                # Hàng tổng cộng
                footer_row_idx = num_records + 1
                total_duration_str = ""
                if total_months_sum > 0:
                    years = total_months_sum // 12
                    months = total_months_sum % 12
                    breakdown = []
                    if years > 0: breakdown.append(f"{years} năm")
                    if months > 0: breakdown.append(f"{months} tháng")
                    total_duration_str = f"{total_months_sum} tháng" + (f" ({' '.join(breakdown)})" if breakdown else "")
                
                cell_start = table.cell(footer_row_idx, 0)
                cell_end = table.cell(footer_row_idx, 4)
                merged_cell = cell_start.merge(cell_end)
                merged_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p_foot1 = merged_cell.paragraphs[0]
                p_foot1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_foot1 = p_foot1.add_run("Tổng thời gian")
                r_foot1.font.name = 'Arial'
                r_foot1.font.size = Pt(10)
                r_foot1.font.bold = True
                
                cell_last = table.cell(footer_row_idx, 5)
                cell_last.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p_foot2 = cell_last.paragraphs[0]
                p_foot2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_foot2 = p_foot2.add_run(total_duration_str)
                r_foot2.font.name = 'Arial'
                r_foot2.font.size = Pt(10)
                r_foot2.font.bold = True

            # ================== TRƯỜNG HỢP 2: TRONG THỐNG KÊ CÓ TỪ 02 NGƯỜI TRỞ LÊN ==================
            else:
                # Khung Cơ quan chủ quản và Quốc hiệu Tiêu ngữ (Bảng ẩn viền)
                header_table = doc.add_table(rows=1, cols=2)
                header_table.autofit = False
                header_table.columns[0].width = Inches(3.0)
                header_table.columns[1].width = Inches(4.5)
                
                cell_left = header_table.cell(0, 0)
                p_left = cell_left.paragraphs[0]
                p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_l1 = p_left.add_run("CÔNG AN TỈNH CAO BẰNG\n")
                r_l1.font.name = 'Arial'
                r_l1.font.size = Pt(11)
                r_l2 = p_left.add_run("PHÒNG TỔ CHỨC CÁN BỘ")
                r_l2.font.name = 'Arial'
                r_l2.font.size = Pt(11)
                r_l2.font.bold = True
                r_l2.font.underline = True
                
                cell_right = header_table.cell(0, 1)
                p_right = cell_right.paragraphs[0]
                p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_r1 = p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
                r_r1.font.name = 'Arial'
                r_r1.font.size = Pt(11)
                r_r1.font.bold = True
                r_r2 = p_right.add_run("Độc lập – Tự do – Hạnh phúc")
                r_r2.font.name = 'Arial'
                r_r2.font.size = Pt(11)
                r_r2.font.bold = True
                r_r2.font.underline = True
                
                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_before = Pt(12)
                
                # Địa danh, ngày tháng
                p_date = doc.add_paragraph()
                p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                current_date_str = f"Cao Bằng, ngày {datetime.now().day} tháng {datetime.now().month} năm {datetime.now().year}"
                r_date = p_date.add_run(current_date_str)
                r_date.font.name = 'Arial'
                r_date.font.size = Pt(11)
                r_date.font.italic = True
                
                p_space2 = doc.add_paragraph()
                p_space2.paragraph_format.space_before = Pt(18)
                
                # Tiêu đề danh sách tổng hợp
                p_title = doc.add_paragraph()
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_title = p_title.add_run("DANH SÁCH TỔNG HỢP QUÁ TRÌNH HƯỞNG PHỤ CẤP")
                r_title.font.name = 'Arial'
                r_title.font.size = Pt(14)
                r_title.font.bold = True
                
                p_space3 = doc.add_paragraph()
                p_space3.paragraph_format.space_before = Pt(18)
                
                # Tính tổng số hàng của tất cả bản ghi của mọi người để khởi tạo một bảng lớn duy nhất
                total_rows_needed = sum(len(emp["all_records"]) for _, emp in self.stats_filtered_items)
                
                table = doc.add_table(rows=total_rows_needed + 1, cols=8, style='Table Grid')
                table.autofit = False
                
                # Định kích thước độ rộng của 8 cột tích hợp
                col_widths = [Inches(0.4), Inches(1.4), Inches(0.8), Inches(1.3), Inches(0.9), Inches(0.9), Inches(0.9), Inches(0.8)]
                for row in table.rows:
                    for idx, width in enumerate(col_widths):
                        row.cells[idx].width = width
                
                # Tạo hàng tiêu đề cho bảng tổng hợp
                headers = ["STT", "Họ và tên", "Số hiệu", "Đơn vị công tác", "TG Bắt đầu", "TG Kết thúc", "Lý do kết thúc", "Thời gian"]
                for idx, h_text in enumerate(headers):
                    cell = table.cell(0, idx)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run(h_text)
                    r.font.name = 'Arial'
                    r.font.size = Pt(10)
                    r.font.bold = True
                
                # Đổ dữ liệu tuần tự và tự động gộp dòng (Merge) thông tin định danh cá nhân
                global_stt = 1
                row_idx = 1
                for _, emp in self.stats_filtered_items:
                    sh_str = emp['cand_id'] if emp['cand_id'] != "Chưa cập nhật" else ""
                    num_records = len(emp["all_records"])
                    start_row = row_idx  # Đánh dấu dòng bắt đầu của nhân sự này
                    
                    for i, rec in enumerate(emp["all_records"]):
                        unit = rec.get("unit", "")
                        if unit == "Chưa cập nhật": unit = ""
                        start = rec.get("start_date", "")
                        if start == "Chưa cập nhật": start = ""
                        end = rec.get("end_date", "")
                        if end == "Chưa cập nhật": end = ""
                        
                        reason = rec["full_row"].get("Lý do kết thúc", rec["full_row"].get("Lý do", ""))
                        if reason == "Chưa cập nhật": reason = ""
                            
                        duration = rec["full_row"].get("Tổng số tháng", rec["full_row"].get("Thời gian", ""))
                        if duration == "Chưa cập nhật": duration = ""
                        if str(duration).isdigit():
                            duration = f"{duration} tháng"
                        
                        vals = [str(global_stt), emp['display_name'], sh_str, str(unit), str(start), str(end), str(reason), str(duration)]
                        alignments = [
                            WD_ALIGN_PARAGRAPH.CENTER,
                            WD_ALIGN_PARAGRAPH.LEFT,
                            WD_ALIGN_PARAGRAPH.CENTER,
                            WD_ALIGN_PARAGRAPH.LEFT,
                            WD_ALIGN_PARAGRAPH.CENTER,
                            WD_ALIGN_PARAGRAPH.CENTER,
                            WD_ALIGN_PARAGRAPH.LEFT,
                            WD_ALIGN_PARAGRAPH.CENTER
                        ]
                        
                        for col_idx, val in enumerate(vals):
                            # Nếu là dòng thứ 2 trở đi của cùng 1 người, bỏ qua các cột STT, Họ tên, Số hiệu để tránh lặp chữ sau khi gộp ô
                            if i > 0 and col_idx in [0, 1, 2]:
                                continue
                                
                            cell = table.cell(row_idx, col_idx)
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            p = cell.paragraphs[0]
                            p.alignment = alignments[col_idx]
                            r = p.add_run(val)
                            r.font.name = 'Arial'
                            r.font.size = Pt(9.5)
                        
                        row_idx += 1
                    
                    # Tiến hành merge (gộp dòng) nếu nhân sự này có từ 2 cập nhật trở lên
                    if num_records > 1:
                        end_row = start_row + num_records - 1
                        for col_idx in [0, 1, 2]:  # Gộp ô cho cột STT, Họ và tên, Số hiệu
                            cell_start = table.cell(start_row, col_idx)
                            cell_end = table.cell(end_row, col_idx)
                            cell_start.merge(cell_end)
                            
                    global_stt += 1  # Chỉ tăng STT khi chuyển sang người tiếp theo

            doc.save(e.path)
            self.page.open(ft.SnackBar(ft.Text(f"Xuất file biểu mẫu Word thành công tại: {e.path}"), bgcolor=ft.Colors.GREEN_600))
        except Exception as ex:
            self.page.open(ft.SnackBar(ft.Text(f"Lỗi hệ thống xuất Word: {str(ex)}"), bgcolor=ft.Colors.RED_600))
        self.page.update()

    def filter_adjustment_list(self, e):
        self.render_adjustment_list(append=False)

    def load_more_adjust_data(self, e):
        self.adjust_page_number += 1
        self.render_adjustment_list(append=True)

    def render_adjustment_list(self, append=False):
        if not append:
            self.adjust_table_container.controls.clear()
            self.adjust_page_number = 1
            
            # SỬ DỤNG LISTVIEW THAY VÌ COLUMN ĐỂ TỐI ƯU RENDER
            self.adjust_grid = ft.ListView(spacing=0, expand=True) 
            
            self.adjust_horizontal_wrapper = ft.Row(
                scroll=ft.ScrollMode.ALWAYS, 
                controls=[
                    # Bọc ListView trong một Container có width đủ rộng để cho phép cuộn ngang
                    ft.Container(content=self.adjust_grid, width=1500, expand=True) 
                ],
                expand=True
            )
            self.adjust_table_container.controls.append(self.adjust_horizontal_wrapper)
        else:
            if len(self.adjust_grid.controls) > 0:
                if getattr(self.adjust_grid.controls[-1], "key", None) == "more_adjust_btn_container":
                    self.adjust_grid.controls.pop()
        
        q = normalize_text(self.adjust_search_field.value)
        
        matching_items = []
        for k, emp in self.master_data.items():
            if not q or (q in normalize_text(emp["display_name"]) or q in normalize_text(emp["cand_id"]) or q in normalize_text(emp["unit"])):
                matching_items.append((k, emp))

        if not matching_items:
            self.adjust_grid.controls.append(
                ft.Padding(padding=20, content=ft.Text("Không tìm thấy dữ liệu nhân sự phù hợp để sửa.", color=ft.Colors.GREY_500, italic=True))
            )
            self.page.update()
            return

        max_name_len = max([len(emp["display_name"]) for _, emp in matching_items] + [10])
        max_cand_len = max([len(str(emp["cand_id"])) for _, emp in matching_items] + [8])
        max_unit_len = 10
        for _, emp in matching_items:
            for rec in emp["all_records"]:
                max_unit_len = max(max_unit_len, len(str(rec["unit"])))

        w_name = min(180, max(100, int(max_name_len * 6.5) + 15))
        w_yob = 70      
        w_cand = max(80, int(max_cand_len * 6.5) + 15)
        w_unit = min(250, max(120, int(max_unit_len * 6.5) + 15))
        w_rate = 75     
        w_start = 85    
        w_end = 85      
        w_extra = 130 # Độ rộng tiêu chuẩn cho các cột ghi chú/mở rộng

        row_height = 38 
        header_height = 35

        grid_rows = []
        
        if not append:
            header_controls = [
                ft.Container(content=ft.Text("Họ tên", weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_900, size=12), width=w_name, height=header_height, bgcolor=ft.Colors.TEAL_50, alignment=ft.alignment.center, border=ft.border.all(0.5, ft.Colors.GREY_300)),
                ft.Container(content=ft.Text("Năm sinh", weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_900, size=12), width=w_yob, height=header_height, bgcolor=ft.Colors.TEAL_50, alignment=ft.alignment.center, border=ft.border.all(0.5, ft.Colors.GREY_300)),
                ft.Container(content=ft.Text("Số hiệu", weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_900, size=12), width=w_cand, height=header_height, bgcolor=ft.Colors.TEAL_50, alignment=ft.alignment.center, border=ft.border.all(0.5, ft.Colors.GREY_300)),
                ft.Container(content=ft.Text("Đơn vị", weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_900, size=12), width=w_unit, height=header_height, bgcolor=ft.Colors.TEAL_50, alignment=ft.alignment.center, border=ft.border.all(0.5, ft.Colors.GREY_300)),
                ft.Container(content=ft.Text("Mức", weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_900, size=12), width=w_rate, height=header_height, bgcolor=ft.Colors.TEAL_50, alignment=ft.alignment.center, border=ft.border.all(0.5, ft.Colors.GREY_300)),
                ft.Container(content=ft.Text("Bắt đầu", weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_900, size=12), width=w_start, height=header_height, bgcolor=ft.Colors.TEAL_50, alignment=ft.alignment.center, border=ft.border.all(0.5, ft.Colors.GREY_300)),
                ft.Container(content=ft.Text("Kết thúc", weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_900, size=12), width=w_end, height=header_height, bgcolor=ft.Colors.TEAL_50, alignment=ft.alignment.center, border=ft.border.all(0.5, ft.Colors.GREY_300)),
            ]
            
            # Sinh Header động cho các cột dư ra (Ghi chú, lâu năm...)
            for ext_header in self.extra_headers:
                header_controls.append(
                    ft.Container(content=ft.Text(ext_header, weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_900, size=12), width=w_extra, height=header_height, bgcolor=ft.Colors.TEAL_50, alignment=ft.alignment.center, border=ft.border.all(0.5, ft.Colors.GREY_300))
                )

            grid_rows.append(ft.Row(header_controls, spacing=0))

        start_idx = (self.adjust_page_number - 1) * self.per_page if append else 0
        end_idx = self.adjust_page_number * self.per_page
        page_items = matching_items[start_idx:end_idx]

        for emp_key, emp in page_items:
            num_records = len(emp["all_records"])
            total_emp_height = num_records * row_height

            def make_cell_changer(e_key, r_idx, field_type, extra_key=None):
                def cell_on_change(e):
                    val = e.control.value
                    target_emp = self.master_data[e_key]
                    target_rec = target_emp["all_records"][r_idx]
                    if field_type == "name": target_emp["display_name"] = val
                    elif field_type == "yob": target_emp["yob"] = val if val.strip() else "Chưa cập nhật"
                    elif field_type == "cand_id": target_emp["cand_id"] = val
                    elif field_type == "unit": 
                        target_rec["unit"] = val
                        target_emp["unit"] = val 
                    elif field_type == "rate": target_rec["rate"] = val
                    elif field_type == "start":
                        target_rec["start_date"] = val
                        c_code, b_desc, c_flet = self.calculate_months_left(val, target_rec["end_date"])
                        target_rec.update({"status_color_code": c_code, "status_badge_desc": b_desc, "status_color_flet": c_flet})
                    elif field_type == "end":
                        target_rec["end_date"] = val
                        c_code, b_desc, c_flet = self.calculate_months_left(target_rec["start_date"], val)
                        target_rec.update({"status_color_code": c_code, "status_badge_desc": b_desc, "status_color_flet": c_flet})
                    elif field_type == "extra" and extra_key:
                        target_rec["full_row"][extra_key] = val
                return cell_on_change

            yob_display = emp["yob"] if emp["yob"] != "Chưa cập nhật" else ""
            cell_tf_padding = ft.padding.symmetric(horizontal=4, vertical=2)

            name_tf = ft.TextField(value=emp["display_name"], border=ft.InputBorder.NONE, dense=True, text_size=12, text_align=ft.TextAlign.CENTER, content_padding=cell_tf_padding, on_change=make_cell_changer(emp_key, 0, "name"))
            yob_tf = ft.TextField(value=yob_display, border=ft.InputBorder.NONE, dense=True, text_size=12, text_align=ft.TextAlign.CENTER, content_padding=cell_tf_padding, on_change=make_cell_changer(emp_key, 0, "yob"))
            cand_tf = ft.TextField(value=emp["cand_id"], border=ft.InputBorder.NONE, dense=True, text_size=12, text_align=ft.TextAlign.CENTER, content_padding=cell_tf_padding, on_change=make_cell_changer(emp_key, 0, "cand_id"))

            record_subrows = []
            for idx, rec in enumerate(emp["all_records"]):
                unit_tf = ft.TextField(value=rec["unit"], border=ft.InputBorder.NONE, dense=True, text_size=12, content_padding=cell_tf_padding, on_change=make_cell_changer(emp_key, idx, "unit"))
                rate_tf = ft.TextField(value=rec["rate"], border=ft.InputBorder.NONE, dense=True, text_size=12, text_align=ft.TextAlign.CENTER, content_padding=cell_tf_padding, on_change=make_cell_changer(emp_key, idx, "rate"))
                start_tf = ft.TextField(value=rec["start_date"], border=ft.InputBorder.NONE, dense=True, text_size=12, text_align=ft.TextAlign.CENTER, content_padding=cell_tf_padding, on_change=make_cell_changer(emp_key, idx, "start"))
                end_tf = ft.TextField(value=rec["end_date"], border=ft.InputBorder.NONE, dense=True, text_size=12, text_align=ft.TextAlign.CENTER, content_padding=cell_tf_padding, on_change=make_cell_changer(emp_key, idx, "end"))

                subrow_controls = [
                    ft.Container(content=unit_tf, width=w_unit, height=row_height, alignment=ft.alignment.center_left, border=ft.border.only(bottom=ft.BorderSide(0.5, ft.Colors.GREY_300), right=ft.BorderSide(0.5, ft.Colors.GREY_300))),
                    ft.Container(content=rate_tf, width=w_rate, height=row_height, alignment=ft.alignment.center, border=ft.border.only(bottom=ft.BorderSide(0.5, ft.Colors.GREY_300), right=ft.BorderSide(0.5, ft.Colors.GREY_300))),
                    ft.Container(content=start_tf, width=w_start, height=row_height, alignment=ft.alignment.center, border=ft.border.only(bottom=ft.BorderSide(0.5, ft.Colors.GREY_300), right=ft.BorderSide(0.5, ft.Colors.GREY_300))),
                    ft.Container(content=end_tf, width=w_end, height=row_height, alignment=ft.alignment.center, border=ft.border.only(bottom=ft.BorderSide(0.5, ft.Colors.GREY_300), right=ft.BorderSide(0.5, ft.Colors.GREY_300)))
                ]

                # Sinh TextField động cho các cột ghi chú/mở rộng
                for ext_header in self.extra_headers:
                    ext_val = rec["full_row"].get(ext_header, "Chưa cập nhật")
                    ext_tf = ft.TextField(value=ext_val, border=ft.InputBorder.NONE, dense=True, text_size=12, text_align=ft.TextAlign.LEFT, content_padding=cell_tf_padding, on_change=make_cell_changer(emp_key, idx, "extra", ext_header))
                    subrow_controls.append(
                        ft.Container(content=ext_tf, width=w_extra, height=row_height, alignment=ft.alignment.center_left, border=ft.border.only(bottom=ft.BorderSide(0.5, ft.Colors.GREY_300), right=ft.BorderSide(0.5, ft.Colors.GREY_300)))
                    )

                record_subrows.append(ft.Row(subrow_controls, spacing=0))

            emp_block_row = ft.Row([
                ft.Container(content=name_tf, width=w_name, height=total_emp_height, alignment=ft.alignment.center, border=ft.border.only(bottom=ft.BorderSide(0.5, ft.Colors.GREY_300), right=ft.BorderSide(0.5, ft.Colors.GREY_300))),
                ft.Container(content=yob_tf, width=w_yob, height=total_emp_height, alignment=ft.alignment.center, border=ft.border.only(bottom=ft.BorderSide(0.5, ft.Colors.GREY_300), right=ft.BorderSide(0.5, ft.Colors.GREY_300))),
                ft.Container(content=cand_tf, width=w_cand, height=total_emp_height, alignment=ft.alignment.center, border=ft.border.only(bottom=ft.BorderSide(0.5, ft.Colors.GREY_300), right=ft.BorderSide(0.5, ft.Colors.GREY_300))),
                ft.Column(controls=record_subrows, spacing=0)
            ], spacing=0)

            grid_rows.append(emp_block_row)

        self.adjust_grid.controls.extend(grid_rows)

        if len(matching_items) > end_idx:
            more_btn = ft.Container(
                key="more_adjust_btn_container",
                content=ft.ElevatedButton(
                    f"Tải thêm ({len(matching_items) - end_idx} dòng tiếp theo)...",
                    icon=ft.Icons.GRID_ON,
                    on_click=self.load_more_adjust_data,
                    style=ft.ButtonStyle(bgcolor=ft.Colors.TEAL_50, color=ft.Colors.TEAL_900)
                ),
                alignment=ft.alignment.center, padding=ft.padding.all(10)
            )
            self.adjust_grid.controls.append(more_btn)

        self.page.update()

def main(page: ft.Page):
    AllowanceDashboardApp(page)

if __name__ == "__main__":
    ft.app(target=main)
