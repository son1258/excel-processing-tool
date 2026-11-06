
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime


# ==================================================
# CONFIG
# ==================================================

st.set_page_config(
    page_title="Excel Processing Tool",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Processing Tool")
st.caption(
    "Upload Excel → Customize columns → Clean data → Generate Excel"
)


# ==================================================
# SESSION STATE
# ==================================================

if "generated_file" not in st.session_state:
    st.session_state.generated_file = None

if "generated_name" not in st.session_state:
    st.session_state.generated_name = None

if "file_key" not in st.session_state:
    st.session_state.file_key = None


# ==================================================
# 1. UPLOAD FILE
# ==================================================

st.header("1. Upload file")

uploaded_file = st.file_uploader(
    "Chọn file Excel hoặc CSV",
    type=["xlsx", "xls", "csv"]
)

if not uploaded_file:
    st.info("Vui lòng upload file để bắt đầu.")
    st.stop()


# ==================================================
# 2. READ FILE
# ==================================================

try:
    file_name = uploaded_file.name
    is_csv = file_name.lower().endswith(".csv")

    if is_csv:
        sheets = {
            "CSV": pd.read_csv(
                uploaded_file,
                dtype=object
            )
        }

    else:
        excel_file = pd.ExcelFile(uploaded_file)

        sheets = {
            sheet_name: pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                dtype=object
            )
            for sheet_name in excel_file.sheet_names
        }

except Exception as error:
    st.error(f"Lỗi đọc file: {error}")
    st.stop()


# ==================================================
# 3. SELECT SHEET
# ==================================================

st.header("2. Chọn Sheet")

selected_sheet = st.selectbox(
    "Sheet muốn xử lý",
    list(sheets.keys())
)

df = sheets[selected_sheet].copy()

# Chuẩn hóa tên cột
df.columns = [
    str(column).strip()
    for column in df.columns
]

source_columns = list(df.columns)


# ==================================================
# 4. FILE INFORMATION
# ==================================================

st.header("3. Thông tin file gốc")

col1, col2, col3 = st.columns(3)

col1.metric(
    "Tổng số dòng",
    f"{len(df):,}"
)

col2.metric(
    "Tổng số cột",
    len(source_columns)
)

col3.metric(
    "Tổng số Sheet",
    len(sheets)
)


with st.expander(
    "Xem danh sách tất cả cột nguồn",
    expanded=True
):

    for index, column in enumerate(
        source_columns,
        start=1
    ):
        st.write(f"{index}. **{column}**")


# ==================================================
# 5. ORIGINAL DATA PREVIEW
# ==================================================

st.subheader("Preview dữ liệu gốc")

st.dataframe(
    df.head(10),
    use_container_width=True,
    hide_index=True
)


# ==================================================
# 6. INITIALIZE MAPPING
# ==================================================

file_key = (
    f"{file_name}_{selected_sheet}_"
    f"{len(source_columns)}"
)

if st.session_state.file_key != file_key:

    st.session_state.file_key = file_key

    st.session_state.mapping = pd.DataFrame({
        "Tên cột đầu ra": source_columns,
        "Mapping từ cột nguồn": source_columns
    })

    st.session_state.generated_file = None
    st.session_state.generated_name = None


# ==================================================
# 7. CUSTOMIZE OUTPUT COLUMNS
# ==================================================

st.divider()

st.header("4. Customize file Excel đầu ra")

st.caption(
    "Thêm, xóa, đổi tên và sắp xếp các cột "
    "muốn xuất ra file Excel mới."
)

mapping = st.data_editor(
    st.session_state.mapping,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key=f"mapping_editor_{file_key}",
    column_config={

        "Tên cột đầu ra": st.column_config.TextColumn(
            "Tên cột đầu ra",
            required=True
        ),

        "Mapping từ cột nguồn": st.column_config.SelectboxColumn(
            "Mapping từ cột nguồn",
            options=[
                "-- Bỏ qua --"
            ] + source_columns,
            required=True
        )
    }
)

# Cập nhật mapping
st.session_state.mapping = mapping.copy()


# ==================================================
# 8. VALIDATE MAPPING
# ==================================================

valid_mapping = mapping.copy()

valid_mapping = valid_mapping.dropna(
    subset=[
        "Tên cột đầu ra",
        "Mapping từ cột nguồn"
    ]
)

valid_mapping["Tên cột đầu ra"] = (
    valid_mapping["Tên cột đầu ra"]
    .astype(str)
    .str.strip()
)

valid_mapping = valid_mapping[
    (valid_mapping["Tên cột đầu ra"] != "")
    & (valid_mapping["Mapping từ cột nguồn"] != "-- Bỏ qua --")
    & (valid_mapping["Mapping từ cột nguồn"].isin(source_columns))
]

# Kiểm tra trùng tên cột
duplicate_names = valid_mapping[
    valid_mapping["Tên cột đầu ra"].duplicated(keep=False)
]["Tên cột đầu ra"].tolist()

if duplicate_names:

    st.error(
        "Tên cột đầu ra bị trùng: "
        + ", ".join(set(duplicate_names))
    )

    st.stop()


if valid_mapping.empty:

    st.warning(
        "Bạn chưa chọn cột nào để xuất."
    )

    st.stop()

# ==================================================
# 9. BUILD OUTPUT DATA - ONLY SELECTED COLUMNS
# ==================================================

selected_mapping = valid_mapping[
    valid_mapping["Tên cột đầu ra"].notna()
    & valid_mapping["Mapping từ cột nguồn"].notna()
    & (valid_mapping["Tên cột đầu ra"].str.strip() != "")
    & (valid_mapping["Mapping từ cột nguồn"] != "-- Bỏ qua --")
    & (valid_mapping["Mapping từ cột nguồn"].isin(source_columns))
].copy()

output_df = pd.DataFrame(index=df.index)

# Chỉ lấy những cột đã được chọn ở mục 4
for _, row in selected_mapping.iterrows():

    target_column = row["Tên cột đầu ra"].strip()
    source_column = row["Mapping từ cột nguồn"]

    output_df[target_column] = df[source_column]

# Đảm bảo thứ tự cột giống thứ tự mapping
output_df = output_df[
    selected_mapping["Tên cột đầu ra"].str.strip().tolist()
]


# ==================================================
# 10. DATA CLEANING
# ==================================================

st.divider()

st.header("5. Làm sạch dữ liệu")

st.caption(
    "Có thể chọn nhiều điều kiện làm sạch cùng lúc."
)

clean_options = st.multiselect(

    "Chọn chức năng làm sạch",

    options=[

        "Xóa dòng trống hoàn toàn",

        "Xóa khoảng trắng đầu/cuối",

        "Chuẩn hóa ô chỉ chứa khoảng trắng thành ô trống",

        "Xóa dòng dữ liệu trùng lặp"

    ],

    default=[
        "Xóa dòng trống hoàn toàn"
    ],

    placeholder="Chọn các chức năng muốn áp dụng..."
)


# Giữ lại dữ liệu gốc sau mapping
cleaned_df = output_df.copy()

original_count = len(cleaned_df)


# ==================================================
# 10.1 TRIM WHITESPACE
# ==================================================

if "Xóa khoảng trắng đầu/cuối" in clean_options:

    cleaned_df = cleaned_df.apply(

        lambda column: column.map(

            lambda value: value.strip()

            if isinstance(value, str)

            else value

        )

    )


# ==================================================
# 10.2 EMPTY STRING TO NA
# ==================================================

if (
    "Chuẩn hóa ô chỉ chứa khoảng trắng thành ô trống"
    in clean_options
):

    cleaned_df = cleaned_df.replace(

        to_replace=r"^\s*$",

        value=pd.NA,

        regex=True

    )


# ==================================================
# 10.3 REMOVE COMPLETELY EMPTY ROWS
# ==================================================

if "Xóa dòng trống hoàn toàn" in clean_options:

    empty_rows = cleaned_df.apply(

        lambda column: column.map(

            lambda value:

                pd.isna(value)

                or (

                    isinstance(value, str)

                    and value.strip() == ""

                )

        )

    ).all(axis=1)

    cleaned_df = cleaned_df.loc[
        ~empty_rows
    ]


# ==================================================
# 10.4 REMOVE DUPLICATES
# ==================================================

if "Xóa dòng dữ liệu trùng lặp" in clean_options:

    cleaned_df = cleaned_df.drop_duplicates()


# Reset index
cleaned_df = cleaned_df.reset_index(drop=True)

removed_count = original_count - len(cleaned_df)


# ==================================================
# 11. CLEANING STATISTICS
# ==================================================

st.subheader("Thống kê sau làm sạch")

col1, col2, col3 = st.columns(3)

col1.metric(
    "Dòng ban đầu",
    f"{original_count:,}"
)

col2.metric(
    "Dòng còn lại",
    f"{len(cleaned_df):,}"
)

col3.metric(
    "Dòng đã loại bỏ",
    f"{removed_count:,}"
)


# ==================================================
# 12. OUTPUT PREVIEW
# ==================================================

st.divider()

st.header("6. Preview file Excel đầu ra")

st.write("**Danh sách cột đầu ra:**")

st.write(
    ", ".join(cleaned_df.columns)
)

st.dataframe(
    cleaned_df.head(20),
    use_container_width=True,
    hide_index=True
)

st.caption(
    f"Tổng cộng {len(cleaned_df):,} dòng "
    f"và {len(cleaned_df.columns)} cột."
)


# ==================================================
# 13. EXPORT FUNCTION
# ==================================================

def export_excel(data):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        data.to_excel(
            writer,
            index=False,
            sheet_name="Data"
        )

        worksheet = writer.sheets["Data"]

        # Freeze header
        worksheet.freeze_panes = "A2"

        # Auto filter
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        # Auto column width
        for column_cells in worksheet.columns:

            max_length = max(

                len(str(cell.value or ""))

                for cell in column_cells

            )

            worksheet.column_dimensions[
                column_cells[0].column_letter
            ].width = min(
                max_length + 3,
                40
            )

    output.seek(0)

    return output.getvalue()


# ==================================================
# 14. GENERATE EXCEL
# ==================================================

st.divider()

st.header("7. Generate Excel")

output_name = st.text_input(
    "Tên file Excel đầu ra",
    value=(
        "excel_custom_"
        + datetime.now().strftime("%Y%m%d")
    )
)


if st.button(

    "🚀 TẠO FILE EXCEL",

    type="primary",

    use_container_width=True

):

    if not output_name.strip():

        st.error(
            "Vui lòng nhập tên file."
        )

    elif cleaned_df.empty:

        st.error(
            "Không có dữ liệu để xuất file."
        )

    else:

        try:

            with st.spinner(
                "Đang tạo file Excel..."
            ):

                excel_bytes = export_excel(
                    cleaned_df
                )

                st.session_state.generated_file = (
                    excel_bytes
                )

                st.session_state.generated_name = (
                    output_name.strip()
                )

                st.session_state.generated_signature = (
                    hash(pd.util.hash_pandas_object(
                        cleaned_df,
                        index=True
                    ).values.tobytes())
                )

            st.success(
                "Đã tạo file Excel thành công!"
            )

        except Exception as error:

            st.error(
                f"Lỗi khi tạo file: {error}"
            )


# ==================================================
# 15. DOWNLOAD FILE
# ==================================================

if st.session_state.generated_file:

    st.download_button(

        label="⬇️ TẢI FILE EXCEL",

        data=st.session_state.generated_file,

        file_name=(
            f"{st.session_state.generated_name}.xlsx"
        ),

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        type="primary",

        use_container_width=True

    )
