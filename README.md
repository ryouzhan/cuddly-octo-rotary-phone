# 📦 发货单处理工具 v2.0 (Web 版)

原桌面版 v1.1 升级到 Web 版的完整项目，三个 Tab 功能完整保留。

## ✨ 包含功能

| Tab | 功能 | 输入 | 输出 |
|-----|------|------|------|
| 📊 FBA 仓库映射 | FBA 仓库精准分区 + 表格美化 | 发货单 Excel（含发货单详情 + 装箱信息） | 一个 Excel（详细数据 + 汇总结果） |
| 🏭 AWD 发货单生成 | SKU 智能匹配 + 货值核算 | 备货单 + 商品列表 | 一个 Excel（详细数据 + 汇总结果） |
| 🧮 PCS 计算工具 | PCS 自动计算 + 汇总 | 采购单 + 发货单 | 两个 Excel（发货单_已更新 + 总PCS汇总） |

## 🚀 三种启动方式

### 方式 1：双击 `start.bat`（最简单，Windows 推荐）
双击即可，自动装依赖、启动服务、打开浏览器。

### 方式 2：命令行启动
```bash
pip install -r requirements.txt
streamlit run app.py
```

### 方式 3：部署到云端（让其他人也能用）
免费方案：推 GitHub → [share.streamlit.io](https://share.streamlit.io) 一键部署。

## 📁 项目结构

```
fhd-web/
├── app.py                  # Streamlit 主入口（Web UI）
├── start.bat               # Windows 一键启动脚本
├── requirements.txt        # 依赖清单
├── README.md               # 本文件
├── 原版备份.py             # 你原来的桌面版（备份用）
├── core/
│   ├── __init__.py
│   ├── fba.py              # FBA 业务逻辑
│   ├── awd.py              # AWD 业务逻辑
│   └── pcs.py              # PCS 业务逻辑
├── data/                   # 临时文件（自动清理）
└── logs/                   # 日志
```

## 🔧 系统要求

- Python 3.9+（[下载](https://www.python.org/downloads/)）
- Windows / macOS / Linux 都行
- 内存 2GB+（处理大文件时建议 4GB+）

## ❓ 常见问题

**Q: 启动后浏览器没自动打开？**
A: 手动访问 http://localhost:8501

**Q: 端口 8501 被占用？**
A: 改用 `streamlit run app.py --server.port 8502`

**Q: 处理大文件很慢？**
A: Streamlit 默认限制 200MB 上传，发货单一般不会到。如果到了可以分批处理。

**Q: 处理完 Excel 表格没颜色？**
A: 用 WPS / Office 2016+ 打开，WPS 老版本可能不显示 PatternFill 颜色。

**Q: 想给其他人用？**
A: 三种方案：
1. 部署到 [Streamlit Cloud](https://share.streamlit.io) - 免费
2. 部署到阿里云/腾讯云 - 一年几百块
3. 让他们自己装 Python + 双击 start.bat

## 📝 版本说明

- **v1.1** - 原 tkinter 桌面版（已备份为 `原版备份.py`）
- **v2.0** - Web 版（本项目），代码结构升级，业务逻辑等价
