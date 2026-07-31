"""
Parse a dbt manifest.json (from `dbt compile` or `dbt docs generate`) and
export column-level lineage / SQL mappings to Excel.

Usage:
    python dbt_manifest_to_excel.py path/to/manifest.json [output.xlsx]

    # Only include specific model(s), by name:
    python dbt_manifest_to_excel.py manifest.json out.xlsx --model fct_orders
    python dbt_manifest_to_excel.py manifest.json out.xlsx --model fct_orders,stg_orders
    python dbt_manifest_to_excel.py manifest.json out.xlsx -m fct_orders -m stg_orders

    # Only include model(s) whose manifest path matches a glob/substring:
    python dbt_manifest_to_excel.py manifest.json out.xlsx --path "marts/finance/*"
    python dbt_manifest_to_excel.py manifest.json out.xlsx --path models/staging/

    # One .xlsx per model file instead of a single combined workbook:
    python dbt_manifest_to_excel.py manifest.json --split
    python dbt_manifest_to_excel.py manifest.json my_exports_dir --split
    python dbt_manifest_to_excel.py manifest.json --split --path marts/finance/

With --split, `output_path` is treated as an output *directory* (created if
missing; default: dbt_model_exports/) and one <model_name>.xlsx is written
per matched node, each containing the same four sheets scoped to just that
model. --model/--path filters still apply to select which models get files.

--model and --path can be combined (a node is included if it matches either)
and each can be passed multiple times or as a comma-separated list. With
neither flag, every model in the manifest is included (previous behavior).
Matching against upstream models is still done against the full manifest, so
lineage resolution through unselected upstream CTEs/models still works
correctly - only the *output rows* are restricted to the selection.

Output workbook has four sheets:
    - Columns:        one row per model column (name, type, description, model)
    - Dependencies:   one row per node -> dependency edge (models + macros)
    - Nodes:          one row per node with raw/compiled SQL and metadata
    - Column Lineage: one row per output column per model, with the SQL
                       expression that derives it and the upstream
                       table(s)/column(s) it reads from

The manifest itself does not store column-level lineage, so the
"Column Lineage" sheet is derived by parsing each model's compiled SQL with
sqlglot. This is best-effort: it resolves aliases used in the model's own
FROM/JOIN clauses and, where possible, maps the referenced relation back to
a dbt model/source name using the manifest. It does not recurse through
upstream models to trace a column back to its ultimate origin.

Requires: pandas, openpyxl, sqlglot (`pip install sqlglot`)
"""

import argparse
import fnmatch
import json
import re
import sys
from pathlib import Path

import pandas as pd

try:
    import sqlglot
    from sqlglot import exp
except ImportError:
    sqlglot = None

# Best-effort mapping from dbt adapter_type (manifest metadata) to sqlglot dialect
ADAPTER_TO_DIALECT = {
    "snowflake": "snowflake",
    "bigquery": "bigquery",
    "redshift": "redshift",
    "postgres": "postgres",
    "duckdb": "duckdb",
    "databricks": "databricks",
    "spark": "spark",
    "trino": "trino",
    "presto": "presto",
    "sqlserver": "tsql",
}


def load_manifest(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _flatten_csv(values):
    """Turn a list of CLI values (each possibly comma-separated) into a
    flat list of trimmed strings, or None if nothing was given."""
    if not values:
        return None
    out = []
    for v in values:
        out.extend(p.strip() for p in v.split(",") if p.strip())
    return out or None


def filter_nodes(nodes: dict, model_names=None, path_patterns=None) -> dict:
    """Restrict a manifest nodes dict to those matching --model (exact name,
    case-insensitive) and/or --path (glob pattern or plain substring against
    the node's manifest `path`). A node matches if it satisfies either filter.
    With neither filter set, all nodes are returned unchanged."""
    if not model_names and not path_patterns:
        return nodes

    name_set = {n.lower() for n in model_names} if model_names else set()
    patterns = path_patterns or []

    filtered = {}
    for unique_id, node in nodes.items():
        name = (node.get("name") or "").lower()
        path = node.get("path") or node.get("original_file_path") or ""

        name_match = bool(name_set) and name in name_set
        path_match = any(
            fnmatch.fnmatch(path, pat) or pat.lower() in path.lower()
            for pat in patterns
        )
        if name_match or path_match:
            filtered[unique_id] = node
    return filtered


def extract_rows(manifest: dict, nodes: dict = None):
    """Walk the given nodes (defaults to all of manifest['nodes']) and build
    three lists of dict rows."""
    if nodes is None:
        nodes = manifest.get("nodes", {})

    column_rows = []
    dependency_rows = []
    node_rows = []

    for unique_id, node in nodes.items():
        resource_type = node.get("resource_type")
        name = node.get("name")
        database = node.get("database")
        schema = node.get("schema")
        alias = node.get("alias")
        depends_on = node.get("depends_on", {}) or {}
        dep_nodes = depends_on.get("nodes", []) or []
        dep_macros = depends_on.get("macros", []) or []

        # --- Columns sheet ---
        columns = node.get("columns", {}) or {}
        if columns:
            for col_name, col_meta in columns.items():
                column_rows.append({
                    "unique_id": unique_id,
                    "model": name,
                    "resource_type": resource_type,
                    "database": database,
                    "schema": schema,
                    "column": col_name,
                    "data_type": col_meta.get("data_type"),
                    "description": col_meta.get("description"),
                    "meta": json.dumps(col_meta.get("meta", {})) if col_meta.get("meta") else None,
                })
        else:
            # still record the model even if no column metadata exists
            column_rows.append({
                "unique_id": unique_id,
                "model": name,
                "resource_type": resource_type,
                "database": database,
                "schema": schema,
                "column": None,
                "data_type": None,
                "description": None,
                "meta": None,
            })

        # --- Dependencies sheet ---
        for dep in dep_nodes:
            dependency_rows.append({
                "unique_id": unique_id,
                "model": name,
                "depends_on_type": "node",
                "depends_on": dep,
            })
        for dep in dep_macros:
            dependency_rows.append({
                "unique_id": unique_id,
                "model": name,
                "depends_on_type": "macro",
                "depends_on": dep,
            })

        # --- Nodes sheet (raw SQL / compiled SQL) ---
        node_rows.append({
            "unique_id": unique_id,
            "model": name,
            "resource_type": resource_type,
            "database": database,
            "schema": schema,
            "alias": alias,
            "materialized": (node.get("config", {}) or {}).get("materialized"),
            "path": node.get("path"),
            "depends_on_nodes": ", ".join(dep_nodes) if dep_nodes else None,
            "depends_on_macros": ", ".join(dep_macros) if dep_macros else None,
            "raw_code": node.get("raw_code") or node.get("raw_sql"),
            "compiled_code": node.get("compiled_code") or node.get("compiled_sql"),
        })

    return column_rows, dependency_rows, node_rows


def build_relation_lookup(manifest: dict) -> dict:
    """Map relation identifiers (alias / schema.alias / db.schema.alias, all
    lowercased) to the dbt model/source name, so SQL table references can be
    resolved back to a friendly model name."""
    lookup = {}
    all_nodes = {}
    all_nodes.update(manifest.get("nodes", {}) or {})
    all_nodes.update(manifest.get("sources", {}) or {})

    for node in all_nodes.values():
        alias = node.get("alias") or node.get("identifier") or node.get("name")
        schema = node.get("schema")
        database = node.get("database")
        name = node.get("name")
        if not alias or not name:
            continue
        keys = {alias.lower()}
        if schema:
            keys.add(f"{schema}.{alias}".lower())
        if schema and database:
            keys.add(f"{database}.{schema}.{alias}".lower())
        for key in keys:
            lookup[key] = name
    return lookup


def _resolve_relation(raw_name: str, lookup: dict) -> str:
    """Try progressively shorter suffixes of a dotted/quoted identifier
    against the relation lookup; fall back to the raw name."""
    if not raw_name:
        return raw_name
    cleaned = raw_name.replace('"', "").replace("`", "").lower()
    parts = cleaned.split(".")
    for i in range(len(parts)):
        candidate = ".".join(parts[i:])
        if candidate in lookup:
            return lookup[candidate]
    return raw_name


def _resolve_cte_source(cte_node, cte_map: dict, relation_lookup: dict, depth: int = 0) -> str:
    """Follow a *strictly transparent* `name AS (SELECT * FROM single_table)`
    CTE back to the real relation it selects from (recursing through chained
    CTEs), so lineage rows show the actual upstream model instead of a CTE
    alias. Only unwraps a bare `SELECT * FROM x` with no JOIN/GROUP BY -
    anything that renames, aggregates, or otherwise transforms columns
    (e.g. `SELECT order_id, SUM(amount) ... GROUP BY order_id`) is left as
    the CTE's own name, since resolving further would misattribute a
    computed/aggregated column to a raw upstream column that doesn't
    actually exist there."""
    cte_name = cte_node.alias_or_name
    if depth > 10:
        return cte_name
    inner_select = cte_node.this
    if not isinstance(inner_select, exp.Select):
        return cte_name
    if inner_select.args.get("joins") or inner_select.args.get("group") or inner_select.args.get("group_"):
        return cte_name
    if not (len(inner_select.expressions) == 1 and isinstance(inner_select.expressions[0], exp.Star)):
        return cte_name
    inner_from = _get_from_clause(inner_select)
    if not inner_from:
        return cte_name
    tables = list(inner_from.find_all(exp.Table))
    if len(tables) != 1:
        return cte_name
    t_name = tables[0].name
    if t_name.lower() in cte_map:
        return _resolve_cte_source(cte_map[t_name.lower()], cte_map, relation_lookup, depth + 1)
    return _resolve_relation(t_name, relation_lookup)


def _get_from_clause(select_node):
    return select_node.args.get("from") or select_node.args.get("from_") or select_node.find(exp.From)


def _get_with_clause(select_node):
    return select_node.args.get("with") or select_node.args.get("with_") or select_node.find(exp.With)


def _resolve_table_ref(table_node, cte_map: dict, relation_lookup: dict) -> str:
    name_lower = table_node.name.lower()
    if name_lower in cte_map:
        return _resolve_cte_source(cte_map[name_lower], cte_map, relation_lookup)
    return _resolve_relation(table_node.name, relation_lookup)


def _build_alias_map(select_node, cte_map: dict, relation_lookup: dict) -> dict:
    """alias -> resolved relation name, from a single SELECT's own FROM/JOIN."""
    alias_map = {}
    from_clause = _get_from_clause(select_node)
    if from_clause:
        for t in from_clause.find_all(exp.Table):
            alias_map[t.alias_or_name.lower()] = _resolve_table_ref(t, cte_map, relation_lookup)
    for join in select_node.args.get("joins", []) or []:
        for t in join.find_all(exp.Table):
            alias_map[t.alias_or_name.lower()] = _resolve_table_ref(t, cte_map, relation_lookup)
    return alias_map


def _analyze_projections(select_node, cte_map: dict, relation_lookup: dict, dialect):
    """Return one dict per output column of this SELECT describing the
    expression that derives it and the table(s)/column(s) it references."""
    alias_map = _build_alias_map(select_node, cte_map, relation_lookup)
    single_table = list(alias_map.values())[0] if len(alias_map) == 1 else None

    results = []
    for projection in select_node.expressions:
        if isinstance(projection, exp.Star):
            results.append({
                "column": "*", "expression": "*", "raw_expression": "*",
                "source_tables": ", ".join(sorted(set(alias_map.values()))) or None,
                "source_columns": None, "is_passthrough": True,
                "parse_note": "star projection - expand columns manually",
            })
            continue

        if isinstance(projection, exp.Alias):
            col_name = projection.alias_or_name
            expr_node = projection.this
        else:
            col_name = projection.alias_or_name
            expr_node = projection

        raw_expression_sql = expr_node.sql(dialect=dialect)

        # Build a copy of the expression with every "alias.column" reference
        # rewritten to "real_table.column" (e.g. "c.customer_id" ->
        # "stg_customers.customer_id"), so the logic is readable without
        # having to cross-reference the FROM/JOIN aliases separately.
        resolved_expr_node = expr_node.copy()
        referenced_cols = []
        referenced_tables = set()
        for col in resolved_expr_node.find_all(exp.Column):
            tbl_alias = col.table
            real_table = alias_map.get(tbl_alias.lower(), tbl_alias) if tbl_alias else single_table
            if real_table:
                referenced_tables.add(real_table)
                referenced_cols.append(f"{real_table}.{col.name}")
                col.set("table", exp.to_identifier(real_table))
            else:
                referenced_cols.append(col.name)
        resolved_expression_sql = resolved_expr_node.sql(dialect=dialect)

        results.append({
            "column": col_name,
            "expression": resolved_expression_sql,
            "raw_expression": raw_expression_sql,
            "source_tables": ", ".join(sorted(referenced_tables)) or None,
            "source_columns": ", ".join(referenced_cols) or None,
            "is_passthrough": isinstance(expr_node, exp.Column),
            "parse_note": None,
        })
    return results


def extract_column_lineage(manifest: dict, nodes: dict = None):
    """For every model node in `nodes` (defaults to all of manifest['nodes'])
    with compiled SQL, produce one row per output column per "stage" of the
    query: each named CTE is analyzed as its own stage (so multi-step
    transformations are fully visible), plus the final output stage. A model
    that ends in the common `select * from final` pattern has that
    pass-through resolved so the `final` CTE's real columns are reported as
    the model's output rather than a bare `*`.

    Note: relation resolution (build_relation_lookup) always uses the full
    manifest, not just `nodes`, so lineage through upstream models that were
    filtered out of the output is still resolved correctly - only which
    models get *rows* in the output is restricted by `nodes`."""
    if sqlglot is None:
        print("sqlglot not installed - skipping Column Lineage sheet "
              "(pip install sqlglot)", file=sys.stderr)
        return []

    dialect = ADAPTER_TO_DIALECT.get(
        (manifest.get("metadata", {}) or {}).get("adapter_type")
    )
    relation_lookup = build_relation_lookup(manifest)
    if nodes is None:
        nodes = manifest.get("nodes", {})
    rows = []

    for unique_id, node in nodes.items():
        if node.get("resource_type") != "model":
            continue
        model_name = node.get("name")
        sql = node.get("compiled_code") or node.get("compiled_sql")
        if not sql or not sql.strip():
            rows.append({
                "model": model_name, "unique_id": unique_id, "stage": None,
                "is_final_output": None, "column": None, "expression": None, "raw_expression": None,
                "source_tables": None, "source_columns": None, "is_passthrough": None,
                "parse_note": "no compiled SQL available (run `dbt compile` first)",
            })
            continue

        try:
            parsed = sqlglot.parse_one(sql, read=dialect)
            select = parsed if isinstance(parsed, exp.Select) else parsed.find(exp.Select)
            if select is None:
                raise ValueError("no SELECT statement found")
        except Exception as e:
            rows.append({
                "model": model_name, "unique_id": unique_id, "stage": None,
                "is_final_output": None, "column": None, "expression": None, "raw_expression": None,
                "source_tables": None, "source_columns": None, "is_passthrough": None,
                "parse_note": f"parse error: {e}",
            })
            continue

        cte_map = {}
        with_clause = _get_with_clause(select)
        if with_clause:
            cte_map = {c.alias_or_name.lower(): c for c in with_clause.expressions}

        # Detect the common "select * from <final_cte>" pass-through pattern
        # so we report that CTE's real columns as the model's output instead
        # of a bare "*".
        pure_passthrough_target = None
        exprs = select.expressions
        if len(exprs) == 1 and isinstance(exprs[0], exp.Star) and not (select.args.get("joins")):
            from_clause = _get_from_clause(select)
            if from_clause:
                tables = list(from_clause.find_all(exp.Table))
                if len(tables) == 1 and tables[0].name.lower() in cte_map:
                    pure_passthrough_target = tables[0].name.lower()

        # stage_label -> (select_node, is_final_output)
        stages = []
        for cte_name, cte_node in cte_map.items():
            if isinstance(cte_node.this, exp.Select):
                is_output = cte_name == pure_passthrough_target
                stages.append((cte_name, cte_node.this, is_output))
        if pure_passthrough_target is None:
            stages.append(("(final select)", select, True))

        for stage_label, stage_select, is_output in stages:
            for proj in _analyze_projections(stage_select, cte_map, relation_lookup, dialect):
                rows.append({
                    "model": model_name,
                    "unique_id": unique_id,
                    "stage": stage_label,
                    "is_final_output": is_output,
                    **proj,
                })

    return rows


def write_excel(column_rows, dependency_rows, node_rows, lineage_rows, output_path: str):
    columns_df = pd.DataFrame(column_rows)
    dependencies_df = pd.DataFrame(dependency_rows)
    nodes_df = pd.DataFrame(node_rows)
    lineage_df = pd.DataFrame(lineage_rows)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        columns_df.to_excel(writer, sheet_name="Columns", index=False)
        dependencies_df.to_excel(writer, sheet_name="Dependencies", index=False)
        nodes_df.to_excel(writer, sheet_name="Nodes", index=False)
        lineage_df.to_excel(writer, sheet_name="Column Lineage", index=False)

    return columns_df, dependencies_df, nodes_df, lineage_df


DEFAULT_OUTPUT_FILE = "dbt_manifest_mapping.xlsx"
DEFAULT_SPLIT_DIR = "dbt_model_exports"


def _safe_filename(name: str) -> str:
    """Sanitize a model name into a safe filename component."""
    cleaned = re.sub(r"[^A-Za-z0-9_.-]+", "_", name or "").strip("_")
    return cleaned or "model"


def parse_args():
    parser = argparse.ArgumentParser(
        description="Parse a dbt manifest.json and export column-level lineage / SQL mappings to Excel."
    )
    parser.add_argument("manifest_path", help="Path to manifest.json (from `dbt compile` or `dbt docs generate`)")
    parser.add_argument(
        "output_path", nargs="?", default=None,
        help=f"Output .xlsx path (default: {DEFAULT_OUTPUT_FILE}), or with --split, "
             f"the output directory (default: {DEFAULT_SPLIT_DIR}/)",
    )
    parser.add_argument(
        "-m", "--model", action="append", default=None, metavar="NAME",
        help="Only include this model by name (exact, case-insensitive). "
             "Repeatable, or comma-separated, e.g. -m fct_orders,stg_orders",
    )
    parser.add_argument(
        "-p", "--path", action="append", default=None, metavar="PATTERN",
        help="Only include models whose manifest path matches this glob "
             "pattern or substring (e.g. 'marts/finance/*' or 'staging/'). "
             "Repeatable, or comma-separated. A node matches if it satisfies "
             "--model or --path (whichever were given).",
    )
    parser.add_argument(
        "--split", action="store_true",
        help="Write one <model_name>.xlsx per matched model instead of a "
             "single combined workbook. `output_path` is then treated as an "
             "output directory.",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    try:
        manifest = load_manifest(args.manifest_path)
    except FileNotFoundError:
        print(f"Error: manifest file not found: {args.manifest_path}", file=sys.stderr)
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"Error: {args.manifest_path} is not valid JSON ({e})", file=sys.stderr)
        sys.exit(1)

    model_names = _flatten_csv(args.model)
    path_patterns = _flatten_csv(args.path)

    all_nodes = manifest.get("nodes", {})
    selected_nodes = filter_nodes(all_nodes, model_names, path_patterns)

    if (model_names or path_patterns) and not selected_nodes:
        print("No nodes matched the given --model/--path filter(s); nothing to export.", file=sys.stderr)
        sys.exit(1)
    if model_names or path_patterns:
        matched_names = sorted({n.get("name") for n in selected_nodes.values()})
        print(f"Filter matched {len(selected_nodes)} node(s): {', '.join(matched_names)}")

    if args.split:
        out_dir = Path(args.output_path or DEFAULT_SPLIT_DIR)
        out_dir.mkdir(parents=True, exist_ok=True)

        used_filenames = set()
        files_written = 0
        for unique_id, node in selected_nodes.items():
            model_name = node.get("name") or unique_id
            single_node = {unique_id: node}

            filename = f"{_safe_filename(model_name)}.xlsx"
            if filename in used_filenames:
                # avoid collisions between distinct nodes that sanitize to the same name
                filename = f"{_safe_filename(model_name)}__{_safe_filename(unique_id)}.xlsx"
            used_filenames.add(filename)
            file_path = out_dir / filename

            column_rows, dependency_rows, node_rows = extract_rows(manifest, single_node)
            lineage_rows = extract_column_lineage(manifest, single_node)
            write_excel(column_rows, dependency_rows, node_rows, lineage_rows, str(file_path))

            print(f"  {model_name} -> {file_path}")
            files_written += 1

        print(f"Wrote {files_written} model workbook(s) to {out_dir.resolve()}")
        return

    output_path = args.output_path or DEFAULT_OUTPUT_FILE
    column_rows, dependency_rows, node_rows = extract_rows(manifest, selected_nodes)
    lineage_rows = extract_column_lineage(manifest, selected_nodes)
    columns_df, dependencies_df, nodes_df, lineage_df = write_excel(
        column_rows, dependency_rows, node_rows, lineage_rows, output_path
    )

    print(f"Wrote {len(columns_df)} column rows, {len(dependencies_df)} dependency rows, "
          f"{len(nodes_df)} node rows, {len(lineage_df)} lineage rows -> {Path(output_path).resolve()}")


if __name__ == "__main__":
    main()
