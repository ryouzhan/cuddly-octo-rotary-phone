# -*- coding: utf-8 -*-
"""
发货单处理工具 v2.0 (Web 版)
原 tkinter 桌面工具迁移到 Streamlit
"""
import os
import sys
import tempfile
import shutil
from datetime import datetime
from pathlib import Path

# 让项目根目录可被 import
sys.path.insert(0, str(Path(__file__).resolve().parent))

import streamlit as st
import pandas as pd

from core import fba, awd, pcs

# ===================== 页面配置 =====================
st.set_page_config(
    page_title="发货单处理工具 v2.0",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ===================== 自定义样式 =====================
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
        font-size: 1.1rem;
        font-weight: 600;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .stProgress > div > div > div > div { background-color: #007bff; }
</style>
""", unsafe_allow_html=True)

# ===================== 顶部 =====================
st.title("📦 发货单处理工具 v2.0 (Web 版)")
st.caption("原桌面版 v1.1 升级 | 三个功能 Tab 完整保留 | 数据不出本地处理 | 适用于 FBA / AWD / PCS 计算")

st.divider()

# ===================== 工具函数 =====================
def save_uploaded_file(uploaded_file, suffix=None):
    """把 Streamlit 上传的文件存到临时目录，返回路径"""
    if suffix is None:
        suffix = Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    tmp.write(uploaded_file.read())
    tmp.flush()
    tmp.close()
    return tmp.name


def cleanup(*paths):
    for p in paths:
        try:
            os.unlink(p)
        except Exception:
            pass


def render_log_box():
    """返回一个 (placeholder, logger) 的 logger 会把消息写到 placeholder"""
    box = st.empty()
    msgs = []

    def logger(msg):
        msgs.append(msg)
        # 每次刷新整个 log 区，最多保留最后 200 条
        text = "\n".join([f"> {m}" for m in msgs[-200:]])
        box.code(text, language="bash")

    return box, logger


# ===================== 三个 Tab =====================
tab1, tab2, tab3 = st.tabs([
    "📊 FBA 仓库映射",
    "🏭 AWD 发货单生成",
    "🧮 PCS 计算工具",
])

# =================================================================
# Tab 1 - FBA 仓库精准映射 + Excel 美化
# =================================================================
with tab1:
    st.header("📊 FBA 仓库精准映射与表格美化")
    st.markdown("""
    - 自动识别发货单中的 FBA 仓库代码 → 精准归类到 **美东 / 美中 / 美西**
    - 合并 `发货单详情` + `装箱信息` 两个 Sheet
    - 输出包含 `详细数据` + `汇总结果` 两个 Sheet
    - 自动按 **货件编号** 给行底色，便于核对
    """)

    fba_in = st.file_uploader(
        "① 选择发货单 Excel（需包含 `发货单详情` 和 `装箱信息` 两个 Sheet）",
        type=["xlsx", "xls"],
        key="fba_in",
    )

    if fba_in:
        st.info(f"已选择：**{fba_in.name}**  ({fba_in.size / 1024:.1f} KB)")

        if st.button("▶ 开始处理", type="primary", key="fba_btn", use_container_width=True):
            log_box, logger = render_log_box()
            progress = st.progress(0.0, text="准备中...")

            tmp_in = save_uploaded_file(fba_in)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            tmp_out = os.path.join(tempfile.gettempdir(), f"fba_result_{ts}.xlsx")

            try:
                def cb(p, msg):
                    progress.progress(p / 100.0, text=f"{msg} ({p}%)")

                fba.process(tmp_in, tmp_out, log=logger, progress=cb)
                progress.progress(1.0, text="✅ 处理完成")

                with open(tmp_out, "rb") as f:
                    st.download_button(
                        label="📥 下载处理结果",
                        data=f.read(),
                        file_name=f"FBA_{Path(fba_in.name).stem}_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                st.success("处理完成！点击上方按钮下载结果。")
            except Exception as e:
                st.error(f"处理失败：{e}")
            finally:
                cleanup(tmp_in, tmp_out)

# =================================================================
# Tab 2 - AWD 智能生成发货单 + 货值核算
# =================================================================
with tab2:
    st.header("🏭 AWD 发货单智能生成")
    st.markdown("""
    - 备货单 + 商品列表 → 智能 SKU 映射（**支持 ZF- 前缀穿透**）
    - 自动按工作表内容识别 Commodities 数据源
    - 计算 **箱规 / 体积 / 体积重**，并自动核算 **总货值（CNY）**
    """)

    col1, col2 = st.columns(2)
    with col1:
        awd_file = st.file_uploader(
            "① 备货单表格",
            type=["xlsx", "xls", "csv"],
            key="awd_in1",
        )
    with col2:
        comm_file = st.file_uploader(
            "② 商品列表",
            type=["xlsx", "xls", "csv"],
            key="awd_in2",
        )

    if awd_file and comm_file:
        st.info(f"备货单：**{awd_file.name}** | 商品列表：**{comm_file.name}**")

        if st.button("▶ 开始执行", type="primary", key="awd_btn", use_container_width=True):
            log_box, logger = render_log_box()
            progress = st.progress(0.0, text="准备中...")

            tmp_awd = save_uploaded_file(awd_file)
            tmp_comm = save_uploaded_file(comm_file)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            tmp_out = os.path.join(tempfile.gettempdir(), f"awd_result_{ts}.xlsx")

            try:
                def cb(p, msg):
                    progress.progress(p / 100.0, text=f"{msg} ({p}%)")

                awd.process(tmp_awd, tmp_comm, tmp_out, log=logger, progress=cb)
                progress.progress(1.0, text="✅ 处理完成")

                with open(tmp_out, "rb") as f:
                    st.download_button(
                        label="📥 下载处理结果",
                        data=f.read(),
                        file_name=f"AWD_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                st.success("处理完成！点击上方按钮下载结果。")
            except Exception as e:
                st.error(f"处理失败：{e}")
            finally:
                cleanup(tmp_awd, tmp_comm, tmp_out)

# =================================================================
# Tab 3 - PCS 计算工具
# =================================================================
with tab3:
    st.header("🧮 PCS 计算工具")
    st.markdown("""
    - 采购单（含 `SKU / 中文品名 / PCS`） + 发货单（含 `SKU / 发货量 / 货件编号`）
    - 自动匹配并计算 **总PCS = 发货量 × 单PCS**
    - 输出：
        1. `发货单_已更新.xlsx` —— 原发货单 + 总PCS
        2. `货件编号_总PCS汇总表.xlsx` —— 按货件编号 + 品名汇总
    """)

    col1, col2 = st.columns(2)
    with col1:
        purchase = st.file_uploader(
            "① 采购单（.xlsx）",
            type=["xlsx"],
            key="pcs_in1",
        )
    with col2:
        delivery = st.file_uploader(
            "② 发货单（.xlsx）",
            type=["xlsx"],
            key="pcs_in2",
        )

    if purchase and delivery:
        st.info(f"采购单：**{purchase.name}** | 发货单：**{delivery.name}**")

        if st.button("▶ 开始计算", type="primary", key="pcs_btn", use_container_width=True):
            log_box, logger = render_log_box()
            progress = st.progress(0.0, text="准备中...")

            tmp_p = save_uploaded_file(purchase, suffix=".xlsx")
            tmp_d = save_uploaded_file(delivery, suffix=".xlsx")
            tmp_out_dir = tempfile.mkdtemp(prefix="pcs_")
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")

            try:
                def cb(p, msg):
                    progress.progress(p / 100.0, text=f"{msg} ({p}%)")

                updated_path, summary_path = pcs.process(
                    tmp_p, tmp_d, tmp_out_dir, log=logger, progress=cb
                )
                progress.progress(1.0, text="✅ 处理完成")

                st.success("处理完成！下方分别下载两个文件：")
                c1, c2 = st.columns(2)
                with open(updated_path, "rb") as f:
                    c1.download_button(
                        "📥 下载 发货单_已更新.xlsx",
                        data=f.read(),
                        file_name=f"发货单_已更新_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with open(summary_path, "rb") as f:
                    c2.download_button(
                        "📥 下载 货件编号_总PCS汇总表.xlsx",
                        data=f.read(),
                        file_name=f"货件编号_总PCS汇总表_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            except Exception as e:
                st.error(f"处理失败：{e}")
            finally:
                cleanup(tmp_p, tmp_d)
                try:
                    shutil.rmtree(tmp_out_dir, ignore_errors=True)
                except Exception:
                    pass

# =================================================================
# 侧边栏 - 帮助
# =================================================================
with st.sidebar:
    st.header("ℹ️ 使用说明")
    st.markdown("""
    ### 三个功能 Tab
    1. **FBA 仓库映射**
       - 上传含 `发货单详情` + `装箱信息` 的 Excel
       - 自动按 379 个 FBA 仓库代码分美东/美中/美西
    2. **AWD 发货单生成**
       - 上传备货单 + 商品列表（SKU 智能匹配）
       - 自动算箱规/体积/货值
    3. **PCS 计算工具**
       - 上传采购单 + 发货单
       - 输出带 PCS 的发货单 + 汇总表

    ### 文件要求
    - 格式：`.xlsx` / `.xls` / `.csv`
    - 大小：< 200MB（系统限制）

    ### 数据安全
    - 全部在 **本机 / 服务器内存** 中处理
    - 处理完立即清理临时文件
    """)

    st.divider()
    st.caption("v2.0 Web 版 | 由原 v1.1 桌面工具升级 | Powered by Streamlit")
