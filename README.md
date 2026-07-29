
# 星海 HQL 装修工具 🚀

## 📖 项目简介
星海 HQL 装修工具用于将 **ETL 流程 XML 文件** 按照执行顺序（jobhops）还原为原始的 **HQL 脚本**。  
它能够自动解析 XML 文件并生成符合星海数据中台规则的脚本，帮助开发者快速完成流程脚本的转换与规范化。

---

## ⚙️ 功能特性
- 将主流程及子流程 XML 文件打包为 zip，自动转换为 HQL 脚本  
- 输出脚本以主流程为单位，一只主流程对应一只 HQL 文件  
- 文件命名规则：以主流程名称命名  
- 自动处理表名、函数、临时表等规范化问题  

---

## 🛠️ 处理规则

### 非核心组件
- 时间参数组件 → 忽略  
- 数据到达等待检查组件 (等待条件 SQL / MySQL 元数据库) → 以注释形式体现  
- MySQL 日志写入组件 → 提取 `table_name` 并以注释形式体现  
- MySQL 下游采集触发组件 → 提取 `table_name` 并以注释形式体现  

### 核心 HIVE 业务 SQL
- 移除所有 `set` 语句  
- 所有表名统一改为 **小写**  
- 临时表统一前缀为 `temp_`  
- 无模式名的表 → 自动加上 `${schema}_dev` 前缀  
- 有模式名的表 → 在模式名后加 `_dev`  
- 保留原有注释内容与格式  
- 将 `ifnull` 函数统一替换为 `coalesce`  
- 脚本末尾若无删除临时表语句 → 自动新增删除语句  

---

## 📦 使用方法

```bash
python kettle_to_hql_general.py --zip <xml文件包路径>.zip --out <输出文件夹路径> --schema <模式名称>

星海逻辑模型转换工具 🧩
convert_model_detail.py

📘 项目简介
星海逻辑模型转换工具用于读取能力开放门户导出的
“模型详情-库：xxx-表：yyy.xls” 格式的 Excel 文件，
并自动生成可直接导入星海数据中台的 逻辑模型模板 .xlsx 文件。

生成的文件包含两个 Sheet：

实体表模板

字段模板

适用于数据中台逻辑模型建设、表结构标准化、模型批量转换等场景。

✨ 功能特性
自动解析模型详情 Excel 文件

按表名生成对应的 .xlsx 模型文件

自动生成 实体表模板 与 字段模板 两个 Sheet

支持指定输出目录

支持无参数运行自动扫描桌面文件

依赖简单、部署方便

📦 使用方法
1. 基本用法
bash
python convert_model_detail.py <输入文件路径> [输出目录]
2. 示例
bash
python convert_model_detail.py "C:\Users\xxx\Desktop\模型详情-库：ap_tenant_user7-表：hall_estimate_d.xls"
指定输出目录：

bash
python convert_model_detail.py "C:\Users\xxx\Desktop\模型详情-库：ap_tenant_user7-表：hall_estimate_d.xls" "D:\output"
3. 无参数运行
无参数运行时，工具会自动查找桌面上的：

Code
模型详情*.xls
📂 输出文件结构
生成的 .xlsx 文件以表名命名，例如：

Code
hall_estimate_d.xlsx
包含两个 Sheet：

Sheet 名称	内容说明
实体表模板	表级元数据（表名、业务含义、存储层级等）
字段模板	字段级元数据（字段名、类型、描述、是否主键等）


🔧 依赖安装
工具依赖以下 Python 包：

xlrd

pandas

openpyxl

安装方式：

bash
pip install xlrd pandas openpyxl
