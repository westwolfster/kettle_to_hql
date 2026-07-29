# 星海数据中台工具集

本仓库包含两个用于星海数据中台的辅助工具：

- 星海 HQL 装修工具：`kettle_to_hql_general.py`
- 星海逻辑模型转换工具：`convert_model_detail.py`

---

## 一、星海 HQL 装修工具（kettle_to_hql_general.py）

### 1. 工具简介

星海 HQL 装修工具用于读取能力开放门户导出的 ETL 流程 XML 文件，并根据执行顺序（jobhops）还原为原始的 HQL 脚本。

输入为打包好的 XML 流程 zip 文件（包含主流程及若干子流程 XML），输出为符合星海数据中台脚本编写规则的 HQL 脚本文件。

输出文件包括：

- 主流程 HQL 脚本（每个主流程一个 HQL 文件）
- 依赖/等待数据表集合 JSON 文件

### 2. 使用方法

```bash
python kettle_to_hql_general.py --zip <xml文件包路径>.zip --out <输出文件夹路径> --schema <模式名称>

示例：

python kettle_to_hql_general.py --zip C:\Users\westw\Downloads\20260726.zip --out C:\Users\westw\Downloads\output_20260726 --schema wt_day

3. 处理规则

非核心组件处理

时间参数组件：直接忽略，不做任何处理。

数据到达等待检查组件（等待条件 SQL / MySQL 元数据库部分）：不处理 SQL，在该组件位置以注释形式体现。

MySQL 日志写入组件：不处理 SQL，提取插入语句中的 table_name 字段，以注释形式体现。

MySQL 下游采集触发组件：不处理 SQL，提取插入语句中的 table_name 字段，以注释形式体现。

核心 HIVE 业务 SQL 处理

移除所有 set 语句。

所有表名统一改为小写。

所有临时表统一增加前缀：temp_。

无模式名的表名，统一加上 <模式名称>_dev 作为新的模式名，例如：

原始表名：

TMP_CDMA_CHARGE_TJ0_${month_no}

转换后：

<模式名称>_dev.temp_cdma_charge_tj0_${month_no}

有模式名的表名，在原模式名后增加 _dev。

保留原脚本中的所有注释（格式、位置、内容不变）。

将脚本中出现的 ifnull 函数统一替换为 coalesce。

在流程输出的 SQL 脚本末尾，如果没有删除本流程建立的临时表语句，将自动新增删除所有本流程临时表的语句。

4. 输出文件说明

示例输出目录结构：

output_20260726/
├── main_flow_hql.sql
└── wait_tables.json

文件类型

示例文件名

内容说明

主流程 HQL 脚本

main_flow_hql.sql

按执行顺序还原的主流程 HQL 脚本

依赖/等待数据表集合 JSON 文件

wait_tables.json

流程中涉及的依赖/等待数据表集合

二、星海逻辑模型转换工具（convert_model_detail.py）

1. 工具简介

星海逻辑模型转换工具用于读取能力开放门户导出的：

模型详情-库：xxx-表：yyy.xls

格式的 Excel 文件，并根据内容生成以表名命名的 .xlsx 文件。

生成的 .xlsx 文件包含两个 Sheet：

实体表模板

字段模板

星海数据中台可直接导入该 .xlsx 文件生成逻辑模型。

2. 使用方法

python convert_model_detail.py <输入文件路径> [输出目录]

示例：

python convert_model_detail.py "C:\Users\xxx\Desktop\模型详情-库：ap_tenant_user7-表：hall_estimate_d.xls"
python convert_model_detail.py "C:\Users\xxx\Desktop\模型详情-库：ap_tenant_user7-表：hall_estimate_d.xls" "D:\output"

无参数运行时，工具会自动查找桌面上的：

模型详情*.xls

3. 输出文件结构

示例输出文件：

hall_estimate_d.xlsx

包含两个 Sheet：

Sheet 名称

内容说明

实体表模板

表级元数据定义

字段模板

字段级元数据定义

4. 依赖说明

工具依赖以下 Python 包：

pip install xlrd pandas openpyxl

三、项目信息

语言：Python

适用场景：星海数据中台 ETL 流程脚本转换、逻辑模型生成

作者：westwolfster

License：MIT
