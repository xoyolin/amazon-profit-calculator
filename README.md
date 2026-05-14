# 📊 亚马逊多国财务利润统计系统 (Amazon Profit Calculator)

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-App-red.svg)](https://streamlit.io/)
[![License: AGPL v3](https://img.shields.io/badge/License-AGPL_v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

专为亚马逊卖家打造的**极速、精准、安全**的财务利润核算工具。彻底告别卡顿报错的传统 Excel VBA 宏，只需一键上传「日期范围报告」，即可自动清洗欧洲站复杂乱码、精准剥离税金，并秒级生成带完美排版的财务分析报表。

🌐 **[👉 点击这里，立即在线免费体验！](https://amazon-profit-calculator.streamlit.app/)**

---

## ✨ 核心亮点 (Features)

- 🌍 **全语种 12 国支持**：全面兼容美国、加拿大、英国、德国、法国、意大利、西班牙、瑞典、荷兰、波兰、比利时、爱尔兰站点的报表。
- 🛡️ **智能排雷算法**：独创 **“有效字段密度识别法 (VFD)”**，完美穿透德国站/瑞典站的“幽灵空列”、多行废话说明及欧洲逗号小数点陷阱，告别“找不到 Type 列”的报错。
- ⚙️ **配置与代码分离 (免代码热更新)**：基于 `Mapping.xlsx` 驱动。想修改翻译或新增语种？直接在 Excel 里改，保存即生效，无需触碰任何 Python 代码！
- 🎨 **企业级自动排版**：导出的结果不仅是数据，而是**带灰底表头、精准列宽、百分比格式化**的专业 Excel 报表，直接可以直接拿给老板汇报。

---

## 🚀 快速上手 (How to Use)

只需简单的 3 步，即可算清一本账：

1. **上传亚马逊报告**：在网页中上传从亚马逊后台下载的 `日期范围报告 (CSV格式)`。
2. **填写产品成本**：系统会自动提取报表中的 SKU，点击下载「SKU成本填写模板」，填入对应的产品及头程成本后上传，成本金额的币种要跟统计的站点一致。
3. **一键导出报表**：点击下载，即可获得包含「订单统计表」和「费用汇总表」的精美 Excel。

---

## 💻 开发者本地部署 (Local Deployment)

如果你是开发者，想在本地运行这套开源代码，请按以下步骤操作：

1. 克隆本项目到本地：
   ```bash
   git clone [https://github.com/你的用户名/你的仓库名.git](https://github.com/你的用户名/你的仓库名.git)
   cd 你的仓库名



