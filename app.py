import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook

st.set_page_config(
    page_title="Excel Template Mapper",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Template Mapper")
st.caption(
    "Tải form mẫu → Upload file dữ liệu → Map cột → Làm sạch → Xuất file theo form mẫu"
)

# ==================================================
# HELPERS
# ==================================================

def read_file(uploaded_file):
    """Read an uploaded Excel/CSV file into a dict of DataFrames."""
    file_name = uploaded_file.name.lower()
    uploaded_file.seek(0)

    if file_name.endswith(".csv"):
        return {"CSV": pd.read_csv(uploaded_file, dtype=object)}

    excel_file = pd.ExcelFile(uploaded_file)
    result = {}
    for sheet_name in excel_file.sheet_names:
        result[sheet_name] = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            dtype=object,
            header=None
        )
    return result


def normalize_header(value, index):
    if pd.isna(value) or str(value).strip() == "":
        return f"Cột {index + 1}"
    return str(value).strip()


def build_blank_template(template_headers):
    """Create a clean xlsx template with the selected template headers."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(columns=template_headers).to_excel(
            writer, index=False, sheet_name="Data"
        )
        ws = writer.sheets["Data"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
    output.seek(0)
    return output.getvalue()


def export_excel(data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Data")
        worksheet = writer.sheets["Data"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for column_cells in worksheet.columns:
            max_length = max(
                (len(str(cell.value or "")) for cell in column_cells),
                default=10
            )
            worksheet.column_dimensions[
                column_cells[0].column_letter
            ].width = min(max_length + 3, 40)

    output.seek(0)
    return output.getvalue()


def make_unique_headers(values):
    """Avoid duplicate DataFrame column labels while keeping visible names readable."""
    result = []
    counts = {}
    for index, value in enumerate(values):
        name = normalize_header(value, index)
        counts[name] = counts.get(name, 0) + 1
        if counts[name] > 1:
            name = f"{name} ({counts[name]})"
        result.append(name)
    return result


# ==================================================
# 1. UPLOAD TEMPLATE + DATA
# ==================================================

st.header("1. Upload form mẫu và file dữ liệu")

left, right = st.columns(2)

with left:
    st.subheader("📄 Form mẫu")
    template_file = st.file_uploader(
        "Upload file Excel mẫu (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="template_uploader",
        help="Form mẫu quyết định tên cột và thứ tự cột của file đầu ra."
    )

with right:
    st.subheader("📥 File dữ liệu")
    data_file = st.file_uploader(
        "Upload file dữ liệu (.xlsx / .xls / .csv)",
        type=["xlsx", "xls", "csv"],
        key="data_uploader",
        help="File dữ liệu chứa các cột và dòng cần đưa vào form mẫu."
    )

if not template_file or not data_file:
    st.info("Vui lòng upload đủ cả form mẫu và file dữ liệu để bắt đầu.")
    st.stop()

try:
    template_sheets_raw = read_file(template_file)
    data_sheets_raw = read_file(data_file)
except Exception as error:
    st.error(f"Không đọc được file: {error}")
    st.stop()

# ==================================================
# 2. SELECT SHEETS + HEADER ROWS
# ==================================================

st.divider()
st.header("2. Chọn Sheet và dòng tiêu đề")

c1, c2 = st.columns(2)

with c1:
    template_sheet_name = st.selectbox(
        "Sheet trong form mẫu",
        list(template_sheets_raw.keys()),
        key="template_sheet"
    )
    template_raw = template_sheets_raw[template_sheet_name]
    template_header_row = st.number_input(
        "Dòng chứa tên cột trong form mẫu",
        min_value=1,
        max_value=max(1, len(template_raw)),
        value=1,
        step=1,
        key=f"template_header_{template_sheet_name}"
    )

with c2:
    data_sheet_name = st.selectbox(
        "Sheet trong file dữ liệu",
        list(data_sheets_raw.keys()),
        key="data_sheet"
    )
    data_raw = data_sheets_raw[data_sheet_name]
    data_header_row = st.number_input(
        "Dòng chứa tên cột trong file dữ liệu",
        min_value=1,
        max_value=max(1, len(data_raw)),
        value=1,
        step=1,
        key=f"data_header_{data_sheet_name}"
    )

template_header_values = template_raw.iloc[int(template_header_row) - 1].tolist()
data_header_values = data_raw.iloc[int(data_header_row) - 1].tolist()

template_columns = make_unique_headers(template_header_values)
data_columns = make_unique_headers(data_header_values)

# Build data frames from the selected header rows
template_preview = template_raw.iloc[int(template_header_row):].copy()
template_preview.columns = template_columns
template_preview = template_preview.reset_index(drop=True)

data_df = data_raw.iloc[int(data_header_row):].copy()
data_df.columns = data_columns
data_df = data_df.reset_index(drop=True)

# Drop fully empty rows from imported data only? Keep them until user chooses cleaning.
st.caption(
    f"Form mẫu: **{len(template_columns)} cột** · "
    f"File dữ liệu: **{len(data_df):,} dòng**, **{len(data_columns)} cột**"
)

with st.expander("Xem trước form mẫu", expanded=True):
    st.dataframe(
        pd.DataFrame(columns=template_columns),
        use_container_width=True,
        hide_index=True
    )
    st.write("**Tên cột và thứ tự trong form mẫu:**")
    st.write(" → ".join(template_columns))

with st.expander("Xem trước file dữ liệu"):
    st.dataframe(data_df.head(10), use_container_width=True, hide_index=True)

# Allow downloading a blank version of the selected template's column structure
try:
    blank_template_bytes = build_blank_template(template_columns)
    st.download_button(
        "⬇️ Tải form mẫu trống (.xlsx)",
        data=blank_template_bytes,
        file_name="form_mau_trong.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except Exception as error:
    st.warning(f"Chưa thể tạo form mẫu trống: {error}")

# ==================================================
# 3. MAP TEMPLATE COLUMNS TO DATA COLUMNS
# ==================================================

st.divider()
st.header("3. Map cột form mẫu với cột dữ liệu")
st.caption(
    "Tên cột đầu ra và thứ tự được lấy tự động từ form mẫu. "
    "Bạn chỉ cần chọn cột tương ứng trong file dữ liệu; không cần sửa tên cột bằng tay."
)

mapping_key = (
    f"{template_file.name}_{data_file.name}_"
    f"{template_sheet_name}_{data_sheet_name}_"
    f"{len(template_columns)}_{len(data_columns)}"
)

if st.session_state.get("mapping_key") != mapping_key:
    st.session_state.mapping_key = mapping_key
    st.session_state.column_mapping = {
        template_col: (
            template_col if template_col in data_columns else "-- Không map --"
        )
        for template_col in template_columns
    }

mapping = {}
for start in range(0, len(template_columns), 2):
    cols = st.columns(2)
    for offset, template_col in enumerate(template_columns[start:start + 2]):
        with cols[offset]:
            default_value = st.session_state.column_mapping.get(
                template_col,
                template_col if template_col in data_columns else "-- Không map --"
            )
            options = ["-- Không map --"] + data_columns
            default_index = options.index(default_value) if default_value in options else 0
            selected_source = st.selectbox(
                f"**{template_col}**  ←  File dữ liệu",
                options=options,
                index=default_index,
                key=f"map_{mapping_key}_{template_col}"
            )
            mapping[template_col] = selected_source

st.session_state.column_mapping = mapping

unmapped_columns = [
    col for col, source in mapping.items() if source == "-- Không map --"
]
if unmapped_columns:
    st.warning(
        "Các cột sau chưa được map và sẽ để trống trong file đầu ra: "
        + ", ".join(unmapped_columns)
    )

# ==================================================
# 4. BUILD OUTPUT USING TEMPLATE STRUCTURE
# ==================================================

output_df = pd.DataFrame(index=data_df.index)

for template_col in template_columns:
    source_col = mapping[template_col]
    if source_col == "-- Không map --":
        output_df[template_col] = pd.NA
    else:
        output_df[template_col] = data_df[source_col]

output_df = output_df[template_columns]

# ==================================================
# 5. DATA CLEANING
# ==================================================

st.divider()
st.header("4. Làm sạch dữ liệu")
st.caption("Có thể chọn nhiều điều kiện làm sạch cùng lúc.")

clean_options = st.multiselect(
    "Chọn chức năng làm sạch",
    options=[
        "Xóa dòng trống hoàn toàn",
        "Xóa khoảng trắng đầu/cuối",
        "Chuẩn hóa ô chỉ chứa khoảng trắng thành ô trống",
        "Xóa dòng dữ liệu trùng lặp"
    ],
    default=["Xóa dòng trống hoàn toàn"],
    placeholder="Chọn các chức năng muốn áp dụng..."
)

cleaned_df = output_df.copy()
original_count = len(cleaned_df)

if "Xóa khoảng trắng đầu/cuối" in clean_options:
    cleaned_df = cleaned_df.apply(
        lambda column: column.map(
            lambda value: value.strip() if isinstance(value, str) else value
        )
    )

if "Chuẩn hóa ô chỉ chứa khoảng trắng thành ô trống" in clean_options:
    cleaned_df = cleaned_df.replace(
        to_replace=r"^\s*$",
        value=pd.NA,
        regex=True
    )

if "Xóa dòng trống hoàn toàn" in clean_options:
    empty_rows = cleaned_df.apply(
        lambda column: column.map(
            lambda value: pd.isna(value)
            or (isinstance(value, str) and value.strip() == "")
        )
    ).all(axis=1)
    cleaned_df = cleaned_df.loc[~empty_rows]

if "Xóa dòng dữ liệu trùng lặp" in clean_options:
    cleaned_df = cleaned_df.drop_duplicates()

cleaned_df = cleaned_df.reset_index(drop=True)
removed_count = original_count - len(cleaned_df)

m1, m2, m3 = st.columns(3)
m1.metric("Dòng ban đầu", f"{original_count:,}")
m2.metric("Dòng còn lại", f"{len(cleaned_df):,}")
m3.metric("Dòng đã loại bỏ", f"{removed_count:,}")

# ==================================================
# 6. OUTPUT PREVIEW
# ==================================================

st.divider()
st.header("5. Preview file đầu ra")
st.write("**Cột đầu ra theo đúng form mẫu:**")
st.write(" → ".join(cleaned_df.columns.tolist()))

st.dataframe(
    cleaned_df.head(20),
    use_container_width=True,
    hide_index=True
)

st.caption(
    f"Tổng cộng {len(cleaned_df):,} dòng và {len(cleaned_df.columns)} cột."
)

# ==================================================
# 7. GENERATE + DOWNLOAD
# ==================================================

st.divider()
st.header("6. Tạo file Excel")

output_name = st.text_input(
    "Tên file Excel đầu ra",
    value=f"excel_theo_form_mau_{datetime.now().strftime('%Y%m%d')}"
)

if st.button("🚀 TẠO FILE EXCEL", type="primary", use_container_width=True):
    if not output_name.strip():
        st.error("Vui lòng nhập tên file.")
    elif cleaned_df.empty:
        st.error("Không có dữ liệu để xuất file.")
    else:
        try:
            with st.spinner("Đang tạo file Excel..."):
                st.session_state.generated_file = export_excel(cleaned_df)
                st.session_state.generated_name = output_name.strip()
            st.success("Đã tạo file Excel thành công!")
        except Exception as error:
            st.error(f"Lỗi khi tạo file: {error}")

if st.session_state.get("generated_file"):
    st.download_button(
        label="⬇️ TẢI FILE EXCEL",
        data=st.session_state.generated_file,
        file_name=f"{st.session_state.generated_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
