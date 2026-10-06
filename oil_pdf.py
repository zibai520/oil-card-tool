import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="中石化加油IC卡台账 提取小计")
st.title("中石化加油IC卡台账 提取小计")
upload_file = st.file_uploader("上传PDF", type="pdf")

def safe_float(s):
    if not s:
        return 0.0
    # 清除数字以外多余字符，兼容千分位逗号
    clean_str = re.sub(r"[^\d\.\-]", "", str(s))
    try:
        return float(clean_str)
    except:
        return 0.0

if upload_file:
    all_result = []
    with pdfplumber.open(upload_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                # 过滤行长度不足的脏行
                if len(row) < 7:
                    continue
                cell0 = str(row[0]).strip()
                if "小计" in cell0:
                    card_no = str(row[2]).strip()
                    # 跳过卡号为空的小计行
                    if not card_no:
                        continue
                    qty_str = str(row[6]).strip()
                    qty = safe_float(qty_str)
                    all_result.append({"卡号": card_no, "加油数量": qty})

    df = pd.DataFrame(all_result)
    # ==========增加序号，从1开始，放在第一列==========
    df.insert(0, "序号", range(1, len(df)+1))

    # 自动计算总和与卡号数量（用于网页显示）
    total_qty = df["加油数量"].sum()
    card_count = len(df)

    # ==========构造合计行，放到表格底部==========
    total_row = pd.DataFrame([{
        "序号": "合计",
        "卡号": f"共{card_count}张卡",
        "加油数量": total_qty
    }])
    df_final = pd.concat([df, total_row], ignore_index=True)

    st.subheader("📋卡号汇总结果")
    # hide_index=True 隐藏左侧默认0,1,2索引列
    st.dataframe(df_final, use_container_width=True, hide_index=True)

    # ==========导出Excel：合计的加油数量改成SUM公式==========
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="卡号小计", index=False)
        ws = writer.book["卡号小计"]
        # 明细行：Excel第2行到第(len(df)+1)行；合计行：第(len(df)+2)行
        last_data_row = len(df) + 1          # 最后一条明细所在Excel行号
        total_row_num = len(df) + 2          # 合计行所在Excel行号
        # 加油数量在第3列（C列）
        ws.cell(row=total_row_num, column=3, value=f"=SUM(C2:C{last_data_row})")
    output.seek(0)
    st.download_button("下载Excel汇总表", data=output, file_name="加油卡小计.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
