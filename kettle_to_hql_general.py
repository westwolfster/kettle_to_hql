#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kettle_to_hql.py  (精简通用版)

将自定义 Kettle/Pentaho XML 流程转为清洗后的 Hive HQL。

输入 ZIP 仅需包含：
  - 若干 *.xml 流程定义
  - rebuild_hql_output.json 示例（可选，不参与处理）

输出：
  - 每个主流程一个 .hql 文件（子流程 SQL 内联进主流程）
  - rebuild_hql_output.json（本批次 wait_table / output_table 汇总）

处理规则：
  1.1 时间参数组件 → 忽略
  1.2 等待组件 → 不输出 SQL，写注释 --等待数据表XXX
  1.3 Hive SQL → 清洗（见下）
  1.4/1.5 MySQL 日志/下发 → 不输出 SQL，写注释
  2.1 删除所有 set 语句
  2.2/2.4 表名小写；无 schema 加 ap_tenant_user7_dev.；
         有 schema 则改为 schema_dev.table；无 schema 使用 --schema 指定的基础名 + _dev
  2.3 临时表前缀统一为 temp_
  2.6 原注释保留
  2.7 ifnull → coalesce
  2.8 末尾补充 DROP 本流程临时表
  主流程 JOB 节点内联子流程；纯子流程不单独出文件

不再依赖 table_name.xlsx，不再改写 set_day/set_month 分区谓词。
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
import zipfile
from collections import defaultdict, OrderedDict
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

try:
    from lxml import etree
except ImportError:
    print("请安装 lxml: pip install lxml")
    sys.exit(1)

# ---------------------------------------------------------------------------
# 正则
# ---------------------------------------------------------------------------
SET_STMT_RE = re.compile(r"^\s*set\s+.+?;?\s*$", re.IGNORECASE | re.MULTILINE)
IFNULL_RE = re.compile(r"\bifnull\s*\(", re.IGNORECASE)

SCHEMA_BASE = "ap_tenant_user7"  # 由命令行 --schema 覆盖

TABLE_REF_RE = re.compile(
    r"""(?ix)
    (
        (?:from|join|into|table|overwrite\s+table|
           create\s+(?:temporary\s+)?table(?:\s+if\s+not\s+exists)?|
           drop\s+table(?:\s+if\s+exists)?|
           alter\s+table|
           insert\s+(?:overwrite\s+)?table)
        \s+
    )
    ([`"]?[\w\.\$\{\}]+[`"]?)
    """,
    re.VERBOSE,
)


# ---------------------------------------------------------------------------
# 1. 解析自定义 XML
# ---------------------------------------------------------------------------
def parse_job(xml_path: Path) -> Tuple[str, List[dict], List[dict]]:
    tree = etree.parse(str(xml_path))
    root = tree.getroot()

    job_name = (
        root.findtext(".//baseInfo/name")
        or root.findtext(".//name")
        or xml_path.stem
    )

    entries: List[dict] = []
    for e in root.xpath("//jobentrys/e"):
        eid = e.findtext("id_jobentry") or ""
        name = e.findtext("name") or ""
        code = (e.findtext("code") or "").upper()
        conn_node = e.find(".//connection/e/valueStr")
        conn = (
            conn_node.text.strip()
            if conn_node is not None and conn_node.text
            else ""
        )

        # 长 SQL 可能拆成多个 <e><nr/><valueStr/></e>，按 nr 拼接
        sql = ""
        for container in ("sql", "custom_sql"):
            parts = []
            for e2 in e.xpath(f".//{container}/e"):
                vs = e2.findtext("valueStr") or ""
                if not vs.strip():
                    continue
                nr_s = e2.findtext("nr") or "0"
                try:
                    nr = int(nr_s)
                except ValueError:
                    nr = 0
                parts.append((nr, vs))
            if parts:
                parts.sort(key=lambda x: x[0])
                sql = "".join(p[1] for p in parts)
                break

        entries.append(
            {
                "id": eid,
                "name": name,
                "code": code,
                "connection": conn,
                "sql": sql,
            }
        )

    hops: List[dict] = []
    for h in root.xpath("//jobhops/e"):
        enabled = (h.findtext("enabled") or "Y").upper()
        if enabled not in ("Y", "YES", "TRUE", "1"):
            continue
        hops.append(
            {
                "from": h.findtext("id_from") or "",
                "to": h.findtext("id_to") or "",
            }
        )

    return job_name, entries, hops


def topological_order(entries: List[dict], hops: List[dict]) -> List[dict]:
    id2entry = {e["id"]: e for e in entries if e["id"]}
    graph: Dict[str, List[str]] = defaultdict(list)
    indeg: Dict[str, int] = {eid: 0 for eid in id2entry}
    for h in hops:
        if h["from"] in id2entry and h["to"] in id2entry:
            graph[h["from"]].append(h["to"])
            indeg[h["to"]] = indeg.get(h["to"], 0) + 1

    queue = [eid for eid, d in indeg.items() if d == 0]
    order: List[str] = []
    while queue:
        n = queue.pop(0)
        order.append(n)
        for m in graph[n]:
            indeg[m] -= 1
            if indeg[m] == 0:
                queue.append(m)

    for e in entries:
        if e["id"] not in order:
            order.append(e["id"])

    return [id2entry[eid] for eid in order if eid in id2entry]


# ---------------------------------------------------------------------------
# 2. SQL 改写
# ---------------------------------------------------------------------------
def rewrite_table_name(raw: str) -> str:
    """表名小写；临时表加 temp_ 前缀；统一补 schema。

    - 无模式名：使用全局 SCHEMA_BASE + "_dev"
    - 有模式名：在原模式名后加 "_dev"（已是 _dev 则不重复加）
    """
    name = raw.strip('`"')
    lower = name.lower()
    schema_dev = f"{SCHEMA_BASE}_dev"

    if "." in lower:
        schema, tbl = lower.rsplit(".", 1)
        if schema.endswith("_dev"):
            return f"{schema}.{tbl}"
        return f"{schema}_dev.{tbl}"

    if lower.startswith(("tmp_", "temp_")):
        clean = re.sub(r"^(tmp_|temp_)", "", lower)
        return f"{schema_dev}.temp_{clean}"

    return f"{schema_dev}.{lower}"


def process_hive_sql(sql: str) -> Tuple[str, Set[str]]:
    """返回 (改写后SQL, 输出表集合)。不做 set_day/set_month 改写。"""
    if not sql or not sql.strip():
        return "", set()

    # 2.1 删除 set 语句
    sql = SET_STMT_RE.sub("", sql)
    # 2.7 ifnull → coalesce
    sql = IFNULL_RE.sub("coalesce(", sql)

    output_tables: Set[str] = set()

    def repl_table(m: re.Match) -> str:
        prefix, raw = m.group(1), m.group(2)
        new = rewrite_table_name(raw)
        ctx = prefix.lower()
        if "insert" in ctx and "overwrite" in ctx:
            output_tables.add(new)
        elif "create" in ctx and "table" in ctx:
            if "temp_" not in new.lower() and "tmp_" not in new.lower():
                output_tables.add(new)
        return prefix + new

    sql = TABLE_REF_RE.sub(repl_table, sql)
    return sql.strip(), output_tables


def extract_wait_table(sql: str, entry_name: str = "") -> Optional[str]:
    """从等待 SQL 或节点名中提取被等待的表名。"""
    if sql:
        for pat in [
            r"(?i)table_name\s*=\s*'([^']+)'",
            r"(?i)\btab\s*=\s*'([^']+)'",
            r"(?i)\btable\s*=\s*'([^']+)'",
        ]:
            m = re.search(pat, sql)
            if m:
                return m.group(1).strip()

        m = re.search(r"(?i)from\s+([a-z0-9_\.]+)", sql)
        if m:
            old = m.group(1)
            skip = {
                "kj_data_make_notice",
                "bp_data_sqoop_log",
                "dim_cloud_list_run_log_new",
                "dim_cloud_list_table_new",
            }
            if old.lower() not in skip and old.lower().split(".")[-1] not in skip:
                return old

    if entry_name:
        m = re.search(
            r"(?i)(?:等待|wait)[_ ]*([A-Za-z][A-Za-z0-9_]*\.[A-Za-z][A-Za-z0-9_]*)",
            entry_name,
        )
        if m:
            return m.group(1)
        m = re.search(
            r"(?i)(?:等待|wait)[_ ]*([A-Za-z][A-Za-z0-9_]{3,})",
            entry_name,
        )
        if m:
            return m.group(1)

    return None


def extract_mysql_log_table(sql: str) -> Optional[str]:
    m = re.search(r"(?i)values\s*\(\s*'([^']+)'", sql)
    return m.group(1) if m else None


# ---------------------------------------------------------------------------
# 3. 展开流程（主流程内联子流程）
# ---------------------------------------------------------------------------
def expand_job_lines(
    job_name: str,
    entries: List[dict],
    hops: List[dict],
    all_jobs: Dict[str, Tuple[List[dict], List[dict]]],
    visited: Optional[Set[str]] = None,
) -> Tuple[List[str], Set[str], Set[str], Set[str]]:
    if visited is None:
        visited = set()
    if job_name in visited:
        return [f"-- [循环引用，已跳过: {job_name}]"], set(), set(), set()
    visited = set(visited)
    visited.add(job_name)

    ordered = topological_order(entries, hops)
    lines: List[str] = []
    wait_tables: Set[str] = set()
    output_tables: Set[str] = set()
    created_temps: Set[str] = set()

    for entry in ordered:
        code = entry["code"]
        name = entry["name"]
        sql = entry["sql"] or ""
        conn = entry["connection"]

        if code in ("TIME_PARAM", "SPECIAL", "SUCCESS", "START", "ABORT"):
            continue

        # 1.2 等待组件
        if code == "WAIT_FOR_SQL" or "等待" in name:
            tbl = extract_wait_table(sql, name)
            if tbl:
                wait_tables.add(tbl)
                lines.append(f"--等待数据表{tbl}")
            else:
                lines.append(f"--等待数据表 (无法解析) {name}")
            lines.append("")
            continue

        # 1.4 / 1.5 MySQL 日志 / 下发
        if "MYSQL" in conn.upper() or (
            code == "SQL"
            and (
                "KJ_DATA_MAKE_NOTICE" in sql.upper()
                or "bp_dim_disp_rely_tableinfo_detail" in sql.lower()
            )
        ):
            tbl = extract_mysql_log_table(sql)
            lines.append(f"-- MYSQL日志/下发 : {tbl or name}")
            lines.append("")
            continue

        # 子流程内联
        if code == "JOB":
            lines.append(f"-- ========== 开始内联子流程: {name} ==========")
            if name in all_jobs:
                sub_entries, sub_hops = all_jobs[name]
                sub_lines, sub_wait, sub_out, sub_temps = expand_job_lines(
                    name, sub_entries, sub_hops, all_jobs, visited
                )
                lines.extend(sub_lines)
                wait_tables.update(sub_wait)
                output_tables.update(sub_out)
                created_temps.update(sub_temps)
            else:
                lines.append(f"-- [未找到子流程 XML: {name}]")
            lines.append(f"-- ========== 结束内联子流程: {name} ==========")
            lines.append("")
            continue

        # 1.3 核心 Hive SQL
        if code == "SQL" and "HIVE" in conn.upper():
            rewritten, outs = process_hive_sql(sql)
            if rewritten:
                lines.append(f"-- >>> entry: {name}")
                lines.append(rewritten)
                lines.append("")
                output_tables.update(outs)
                for m in re.finditer(
                    r"(?i)create\s+(?:temporary\s+)?table\s+(?:if\s+not\s+exists\s+)?([`\"]?[\w\.\$\{\}]+[`\"]?)",
                    rewritten,
                ):
                    t = m.group(1).strip('`"')
                    if "temp_" in t.lower() or "tmp_" in t.lower():
                        created_temps.add(t)
            continue

        if sql.strip():
            lines.append(f"-- [unhandled {code}] {name}")
            lines.append("")

    return lines, wait_tables, output_tables, created_temps


def convert_one_job(
    xml_path: Path,
    out_dir: Path,
    all_jobs: Dict[str, Tuple[List[dict], List[dict]]],
    sub_job_names: Set[str],
) -> Optional[dict]:
    job_name, entries, hops = parse_job(xml_path)
    has_job_refs = any(e["code"] == "JOB" for e in entries)

    # 纯子流程不单独输出，由主流程内联
    if job_name in sub_job_names and not has_job_refs:
        print(f"  skip sub-job (inlined by parent): {job_name}")
        return None

    lines, wait_tables, output_tables, created_temps = expand_job_lines(
        job_name, entries, hops, all_jobs
    )

    header = [
        "-- ============================================================",
        f"-- Job : {job_name}",
        f"-- Source: {xml_path.name}",
        "-- ============================================================\n",
    ]
    full_lines = header + lines

    if created_temps:
        full_lines.append("-- ========== 清理本流程临时表 ==========")
        for t in sorted(created_temps):
            full_lines.append(f"DROP TABLE IF EXISTS {t};")
        full_lines.append("")

    safe_name = re.sub(r'[\\/:*?"<>|]', "_", job_name)
    hql_path = out_dir / f"{safe_name}.hql"
    hql_path.write_text("\n".join(full_lines), encoding="utf-8")
    print(f"  wrote {hql_path}  ({len(full_lines)} lines)")

    return {
        "name": safe_name,
        "wait_table": sorted(wait_tables),
        "output_table": sorted(output_tables),
    }


# ---------------------------------------------------------------------------
# 4. 入口
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert Kettle job XML packages into cleaned HQL scripts"
    )
    parser.add_argument("--zip", help="source ZIP（仅含 xml + 可选 json 示例）")
    parser.add_argument("--dir", help="already-extracted folder")
    parser.add_argument("--out", default="./hql_output", help="output directory")
    parser.add_argument(
        "--schema",
        default="ap_tenant_user7",
        help="无模式名的表使用的基础 schema（实际写入为 {schema}_dev），默认 ap_tenant_user7",
    )
    args = parser.parse_args()

    if not args.zip and not args.dir:
        parser.error("either --zip or --dir is required")

    global SCHEMA_BASE
    SCHEMA_BASE = (args.schema or "ap_tenant_user7").strip().rstrip(".")
    print(f"SCHEMA_BASE = {SCHEMA_BASE}  (bare tables -> {SCHEMA_BASE}_dev.*)")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    work_dir = None
    if args.zip:
        work_dir = Path(tempfile.mkdtemp(prefix="kettle_"))
        with zipfile.ZipFile(args.zip, "r") as zf:
            zf.extractall(work_dir)
        for f in list(work_dir.rglob("*")):
            if "#U" in f.name:
                new = re.sub(
                    r"#U([0-9a-fA-F]{4})",
                    lambda m: chr(int(m.group(1), 16)),
                    f.name,
                )
                f.rename(f.with_name(new))
        src = work_dir
    else:
        src = Path(args.dir)

    xml_files = sorted(src.rglob("*.xml"))
    if not xml_files:
        print("ERROR: no XML files found")
        sys.exit(1)

    all_jobs: Dict[str, Tuple[List[dict], List[dict]]] = {}
    xml_by_name: Dict[str, Path] = {}
    sub_job_names: Set[str] = set()

    for xml in xml_files:
        jname, entries, hops = parse_job(xml)
        all_jobs[jname] = (entries, hops)
        xml_by_name[jname] = xml
        for e in entries:
            if e["code"] == "JOB" and e["name"]:
                sub_job_names.add(e["name"])

    results = OrderedDict()
    idx = 1
    for jname, xml in xml_by_name.items():
        print(f"Processing {xml.name} ...")
        try:
            info = convert_one_job(xml, out_dir, all_jobs, sub_job_names)
            if info is not None:
                results[f"etl_{idx}"] = info
                idx += 1
        except Exception as e:
            print(f"  FAILED: {e}")
            import traceback

            traceback.print_exc()

    json_path = out_dir / "rebuild_hql_output.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({"rebuild_hql": results}, f, ensure_ascii=False, indent=4)
    print(f"\nWrote summary → {json_path}")

    if work_dir:
        shutil.rmtree(work_dir, ignore_errors=True)
    print("Done.")


if __name__ == "__main__":
    main()
