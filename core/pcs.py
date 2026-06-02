# -*- coding: utf-8 -*-
"""
PCS 计算工具（从原 ShipmentDataProcessor.process_data 剥离 UI 后的业务逻辑）

输入：采购单 + 发货单
输出：
  1. 发货单_已更新.xlsx     (发货单匹配上 PCS + 总 PCS)
  2. 货件编号_总PCS汇总表.xlsx (按货件编号 + 中文品名汇总总 PCS)
"""
import os
import traceback
import pandas as pd


def process(purchase_path, delivery_path, output_dir, log=None, progress=None):
    """
    PCS 计算主流程
    :param purchase_path: 采购单 xlsx
    :param delivery_path: 发货单 xlsx
    :param output_dir:    输出目录
    :param log:           日志回调
    :param progress:      进度回调 0-100
    :return: (updated_delivery_path, summary_path)
    """
    def emit(msg):
        (log or print)(msg)

    def prog(p, msg=""):
        if progress:
            progress(p, msg)
        if msg:
            emit(msg)

    try:
        prog(10, "正在读取采购单数据...")
        purchase_df = pd.read_excel(purchase_path)
        purchase_df.columns = purchase_df.columns.str.strip()

        prog(30, "正在读取发货单数据...")
        delivery_df = pd.read_excel(delivery_path)
        delivery_df.columns = delivery_df.columns.str.strip()

        required_purchase_cols = ["SKU", "中文品名", "PCS"]
        required_delivery_cols = ["SKU", "发货量", "货件编号"]
        for col in required_purchase_cols:
            if col not in purchase_df.columns:
                raise ValueError(f"采购单缺少必需列：【{col}】\n请检查 Excel 列名是否正确")
        for col in required_delivery_cols:
            if col not in delivery_df.columns:
                raise ValueError(f"发货单缺少必需列：【{col}】\n请检查 Excel 列名是否正确")

        prog(55, "正在匹配 SKU 数据...")
        purchase_mapping = purchase_df[["SKU", "中文品名", "PCS"]].drop_duplicates(subset=["SKU"])
        delivery_updated = pd.merge(
            delivery_df,
            purchase_mapping,
            on="SKU",
            how="left"
        )

        prog(75, "正在计算总 PCS...")
        delivery_updated["PCS"] = pd.to_numeric(delivery_updated["PCS"], errors="coerce").fillna(0)
        delivery_updated["总PCS"] = delivery_updated["发货量"] * delivery_updated["PCS"]

        prog(85, "正在汇总数据...")
        summary_data = delivery_updated.groupby(["货件编号", "中文品名"]).agg({
            "总PCS": "sum"
        }).reset_index()[["货件编号", "中文品名", "总PCS"]]

        prog(95, "正在保存结果文件...")
        os.makedirs(output_dir, exist_ok=True)
        updated_delivery_path = os.path.join(output_dir, "发货单_已更新.xlsx")
        summary_path = os.path.join(output_dir, "货件编号_总PCS汇总表.xlsx")
        delivery_updated.to_excel(updated_delivery_path, index=False)
        summary_data.to_excel(summary_path, index=False)

        prog(100, f"✅ 已生成 2 个文件，保存到：{output_dir}")
        return updated_delivery_path, summary_path
    except Exception as e:
        emit(f"【报错】: {str(e)}")
        emit(traceback.format_exc())
        raise
