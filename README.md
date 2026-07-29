⭐ 星海数据中台工具集
🧩 星海 HQL 装修工具
kettle_to_hql_general.py

📘 项目简介
星海 HQL 装修工具用于读取能力开放门户导出的 ETL 流程 XML 文件，并根据执行顺序（jobhops）还原为原始 HQL 脚本。
输出文件包括：

主流程 HQL 脚本

依赖/等待数据表集合 JSON 文件

🔧 使用方法
bash
python kettle_to_hql_general.py --zip <xml文件包路径>.zip --out <输出文件夹路径> --schema <模式名称>
示例：

bash
python kettle_to_hql_general.py --zip C:\Users\westw\Downloads\20260726.zip --out C:\Users\westw\Downloads\output_20260726 --schema wt_day
📂 输出文件说明
文件类型	示例文件名	内容说明
主流程 HQL 脚本	main_flow_hql.sql	按执行顺序还原的主流程脚本
依赖/等待数据表集合 JSON	wait_tables.json	流程中涉及的依赖表集合


🧩 XML → HQL 转换流程（图片方式）
GitHub README 不支持 Mermaid 渲染，请将流程图导出为 PNG/SVG 后再引用。

markdown
![流程图](docs/flowchart.png)
🍀 星海逻辑模型转换工具
convert_model_detail.py

📘 项目简介
星海逻辑模型转换工具用于读取能力开放门户导出的
《模型详情-库：xxx-表：yyy.xls》 文件，并自动生成可直接导入星海数据中台的逻辑模型模板 .xlsx 文件。

生成的文件包含两个 Sheet：

实体表模板

字段模板

🔧 使用方法
bash
python convert_model_detail.py <输入文件路径> [输出目录]
示例：

bash
python convert_model_detail.py "C:\Users\xxx\Desktop\模型详情-库：ap_tenant_user7-表：hall_estimate_d.xls"
python convert_model_detail.py "C:\Users\xxx\Desktop\模型详情-库：ap_tenant_user7-表：hall_estimate_d.xls" "D:\output"
无参数运行时，工具会自动查找桌面上的：

Code
模型详情*.xls
📂 输出文件结构
生成的 .xlsx 文件以表名命名，例如：

Code
hall_estimate_d.xlsx
包含两个 Sheet：

Sheet 名称	内容说明
实体表模板	表级元数据
字段模板	字段级元数据


📦 依赖安装
bash
pip install xlrd pandas openpyxl
📜 License
MIT License
