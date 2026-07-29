# ✨ 星海数据中台工具集

> 一套专为**星海数据中台**设计的辅助工具，帮助你快速完成 ETL 流程脚本转换与逻辑模型生成。

[![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)]()

---

## 📦 工具一览

本仓库包2只实用工具：

| 工具名称 | 脚本文件 | 主要功能 |
|---------|----------|----------|
| **星海 HQL 装修工具** | `kettle_to_hql_general.py` | 将能力开放门户导出的 Kettle ETL 流程 XML 转换为符合星海规范的 HQL 脚本 |
| **星海逻辑模型转换工具** | `convert_model_detail.py` | 将模型详情 Excel 转换为可直接导入星海数据中台的逻辑模型文件 |

---

## 一、星海 HQL 装修工具

### 工具简介

`kettle_to_hql_general.py` 用于读取能力开放门户导出的 ETL 流程 XML 文件，并根据执行顺序（`jobhops`）还原为原始的 HQL 脚本。

- **输入**：打包好的 XML 流程 zip 文件（包含主流程及若干子流程 XML）、JSON文件模板
- **输出**：符合星海数据中台脚本编写规范的 HQL 脚本文件

**输出文件包括：**

- 主流程 HQL 脚本（每个主流程一个 HQL 文件）
- 依赖 / 等待数据表集合 JSON 文件

### 使用方法

```bash
python kettle_to_hql_general.py --zip <xml文件包路径>.zip --out <输出文件夹路径> --schema <模式名称>
```

示例：

python kettle_to_hql_general.py --zip C:\Users\westw\Downloads\20260726.zip --out C:\Users\westw\Downloads\output_20260726 --schema wt_day
### 处理规则
#### 非核心组件处理
- [ ] 时间参数组件
直接忽略，不做任何处理
- [ ] 数据到达等待检查组件（等待条件 SQL / MySQL 元数据库部分）
不处理 SQL，在该组件位置以注释形式体现
- [ ] MySQL 日志写入组件
不处理 SQL，提取插入语句中的 table_name 字段，以注释形式体现
- [ ] MySQL 下游采集触发组件：不处理 SQL，提取插入语句中的 table_name 字段，以注释形式体现
#### 核心 HIVE 业务 SQL 处理
- 移除 ****~~set 语句~~****
- 所有表名统一改为小写
- 临时表使用统一前缀：***temp_***
- 无模式名的表名，以 ***<模式名称>_dev*** 作为新的模式名
- 有模式名的表名，在原模式名后增加 ***_dev***
- 保留原脚本中的所有注释（格式、位置、内容不变）
- 将脚本中出现的 ~~ifnull~~ 函数统一替换为 ***coalesce***
- 在流程输出的 SQL 脚本末尾，如果没有删除本流程建立的临时表语句，将自动新增删除所有本流程临时表的语句



