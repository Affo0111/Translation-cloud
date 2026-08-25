# Translation Cloud

订单翻译系统（云端版），部署于 Streamlit Cloud。

## 功能

- **A 系统**：按 SKU 翻译模板规则翻译订单定制项。
- **B 系统 / Callie 定制项生成**：根据翻译模板 D/E/F 列生成 callie 定制项（AZ 列）。
- **A 系统规则草稿生成**：输入「翻译前 / 翻译后」定制项，自动生成 A 系统翻译规则草稿。
- **B 系统规则草稿生成**：选择 SKU，根据 E/F 列自动生成 callie 规则草稿。

## 部署

1. 在 [Streamlit Cloud](https://streamlit.io/cloud) 新建 App，指向本仓库。
2. 主文件选择 `translator_streamlit.py`。
3. 部署完成后即可通过 URL 访问。

## 本地模板

- `template_database.json` 为运行时自动收录的本地规则库，**未提交到仓库**。
- 每次使用需在界面手动上传翻译模板 Excel。
