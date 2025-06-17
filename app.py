import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import RGBColor, Pt
from io import BytesIO
from datetime import date
import os
import copy
import re

FIELD_MAP = {
    "No. of assort.:" : "Assortment Breakdown",
    "FOB port / price:" : ("FOB Point", "FOB NB"),
    "Sample send date:" : "=today()",
    "Item No:" : "ITEM#",
    "Description:" : "Item Description"
}

# Normalizing function for flexible matching
def normalize(text):
    return re.sub(r"[\\s:：]+", "", text.strip().lower())

# Create normalized map for comparison
NORMALIZED_MAP = {normalize(k): k for k in FIELD_MAP.keys()}

def resolve_title_key(title):
    key = normalize(title)
    return NORMALIZED_MAP.get(key, None)

def fill_label_table(table, data_row):
    for row in table.rows:
        # Iterate for both left and right columns (1st-2nd and 3rd-4th)
        for ti, vi in [(0, 1), (2, 3)]:  # Left title -> left value and right title -> right value
            if len(row.cells) <= vi:
                continue
            title = row.cells[ti].text.strip()
            target_cell = row.cells[vi]

            matched_key = resolve_title_key(title)
            if matched_key:
                source = FIELD_MAP[matched_key]
                if source == "=today()":
                    value = str(date.today())
                elif isinstance(source, tuple):
                    values = [str(data_row.get(col, "")) for col in source]
                    value = " / ".join(values)
                else:
                    value = str(data_row.get(source, ""))

                # Clear existing content and rewrite the value
                for p in target_cell.paragraphs:
                    target_cell._element.remove(p._element)

                new_paragraph = target_cell.add_paragraph()
                run = new_paragraph.add_run(value)
                run.font.color.rgb = RGBColor(0, 0, 0)
                run.font.size = Pt(10)

                st.write(f"[DEBUG] 标题: {title}, 写入值: {value}, 写入后内容: {target_cell.text}")

def duplicate_table_to_new_section(doc, table):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    new_table = copy.deepcopy(table)
    doc._body._element.append(copy.deepcopy(OxmlElement('w:br')))
    doc._body._element.append(new_table._element)
    return new_table

# Streamlit interface for file upload and download
st.title("📍 标贴填写自动化工具")

uploaded_excel = st.file_uploader("上传 Excel 数据", type=["xlsx"])

if uploaded_excel:
    df = pd.read_excel(uploaded_excel)
    st.success(f"成功读取 {len(df)} 条数据")

    required_cols = set()
    for v in FIELD_MAP.values():
        if isinstance(v, str) and not v.startswith("="):
            required_cols.add(v)
        elif isinstance(v, tuple):
            required_cols.update(v)

    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"Excel 缺少必需列：{', '.join(missing_cols)}")
        st.stop()

    TEMPLATE_PATH = "Kmart Buy Trip Label Template.docx"  # Path for template in Streamlit Cloud
    if not os.path.exists(TEMPLATE_PATH):
        st.error("未找到固定模板文件")
    else:
        doc = Document(TEMPLATE_PATH)
        big_table_template = doc.tables[0]

        template_table = None
        for row in big_table_template.rows:
            for i, cell in enumerate(row.cells):
                if i == 1:
                    continue
                if cell.tables:
                    template_table = cell.tables[0]
                    break
            if template_table:
                break

        if not template_table:
            st.error("未找到标贴模板小表格")
            st.stop()

        all_label_tables = []
        rows_needed = len(df)
        labels_per_table = sum(1 for row in big_table_template.rows for i in [0, 2] if len(row.cells) > i)
        num_full_tables = (rows_needed + labels_per_table - 1) // labels_per_table

        for t in range(num_full_tables):
            new_big_table = copy.deepcopy(big_table_template)
            doc._body._element.append(new_big_table._element)

            for row in new_big_table.rows:
                for i, cell in enumerate(row.cells):
                    if i == 1:
                        continue
                    cell._element.clear_content()
                    new_table = copy.deepcopy(template_table)
                    cell._element.append(new_table._element)
                    all_label_tables.append(new_table)

        st.info(f"总共生成了 {len(all_label_tables)} 个标贴区")

        preview_index = st.number_input(f"Enter preview index (0 to {len(all_label_tables)-1}): ", min_value=0, max_value=len(all_label_tables)-1, value=0)

        if st.button("预览指定标贴"):
            fill_label_table(all_label_tables[preview_index], df.iloc[preview_index % len(df)])
            preview_output = BytesIO()
            doc.save(preview_output)
            st.download_button(f"预览标贴_{preview_index}.docx", data=preview_output.getvalue(), file_name=f"Preview_Label_{preview_index}.docx")

        if st.button("填充所有标贴并保存"):
            for i, (_, row) in enumerate(df.iterrows()):
                if i >= len(all_label_tables):
                    break
                fill_label_table(all_label_tables[i], row)

            output = BytesIO()
            doc.save(output)
            st.download_button("下载填写后的标贴", data=output.getvalue(), file_name="Filled_Labels.docx")
