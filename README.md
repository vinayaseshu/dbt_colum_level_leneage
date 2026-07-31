# dbt Manifest to Excel

Parses a dbt `manifest.json` and exports model metadata, dependencies, and
**column-level SQL lineage** to Excel workbooks. Lineage isn't stored in the
manifest itself, so this tool derives it by parsing each model's compiled
SQL with [`sqlglot`](https://github.com/tobymao/sqlglot).

## Requirements

```
pip install pandas openpyxl sqlglot
```

- **pandas** and **openpyxl** are required for all functionality.
- **sqlglot** is required for the "Column Lineage" sheet only. If it isn't
  installed, the script still runs and produces the other three sheets, with
  a warning printed to stderr.

Your `manifest.json` must come from `dbt compile` or `dbt docs generate` (not
`dbt parse`), since column lineage parsing relies on `compiled_code` /
`compiled_sql` - the fully resolved SQL with `{{ ref() }}` / `{{ source() }}`
macros already substituted with real table names. A manifest without
compiled SQL will still produce the Columns/Dependencies/Nodes sheets, but
each model's lineage rows will show a `parse_note` explaining SQL wasn't
available.

## Quick start

```bash
python dbt_manifest_to_excel.py manifest.json
```

Writes `dbt_manifest_mapping.xlsx` in the current directory, containing every
model/seed/snapshot node in the manifest.

```bash
python dbt_manifest_to_excel.py manifest.json my_output.xlsx
```

Same, but to a specific output path.

## Filtering to specific models

Use `--model`/`-m` to select by exact model name, and/or `--path`/`-p` to
select by a glob pattern or plain substring against the node's manifest
`path`. Both flags are repeatable and accept comma-separated values. A node
is included if it matches **either** filter (not both required).

```bash
# by name
python dbt_manifest_to_excel.py manifest.json out.xlsx --model fct_orders
python dbt_manifest_to_excel.py manifest.json out.xlsx --model fct_orders,stg_orders
python dbt_manifest_to_excel.py manifest.json out.xlsx -m fct_orders -m stg_orders

# by path (glob or substring)
python dbt_manifest_to_excel.py manifest.json out.xlsx --path "marts/finance/*"
python dbt_manifest_to_excel.py manifest.json out.xlsx --path models/staging/
```

With no `--model`/`--path` given, every node in the manifest is included
(default behavior).

Lineage resolution always uses the **full** manifest to resolve upstream
relation names, even when the output rows are restricted by a filter - so a
filtered model's lineage will still correctly show real upstream model names
instead of leaving them unresolved.

## One workbook per model (`--split`)

Add `--split` to write one `<model_name>.xlsx` per matched model instead of a
single combined workbook. `output_path` is then treated as an output
**directory** (created if it doesn't exist).

```bash
python dbt_manifest_to_excel.py manifest.json --split
# -> dbt_model_exports/stg_orders.xlsx, dbt_model_exports/fct_orders.xlsx, ...

python dbt_manifest_to_excel.py manifest.json my_exports_dir --split

# combine with a filter to only split out a subset
python dbt_manifest_to_excel.py manifest.json --split --path marts/finance/
```

Each per-model file contains the same four sheets as the combined workbook,
scoped to just that model. Model names are sanitized for use as filenames
(non-alphanumeric characters become `_`); if two distinct nodes sanitize to
the same filename, the node's `unique_id` is appended to disambiguate.

## Output: the four sheets

Every run (combined or `--split`) produces these sheets:

### Columns

One row per column of every included model/seed/snapshot (or one row with
`column = null` if the node has no column metadata in the manifest).

| field | description |
|---|---|
| `unique_id` | dbt node unique ID |
| `model` | model/node name |
| `resource_type` | `model`, `seed`, `snapshot`, etc. |
| `database`, `schema` | target database/schema |
| `column` | column name |
| `data_type` | from the manifest's column metadata, if present |
| `description` | column description, if documented |
| `meta` | any `meta` config on the column, as JSON |

### Dependencies

One row per `depends_on` edge (both upstream nodes and macros used).

| field | description |
|---|---|
| `unique_id`, `model` | the dependent node |
| `depends_on_type` | `node` or `macro` |
| `depends_on` | the upstream node/macro unique ID |

### Nodes

One row per node with SQL and materialization metadata.

| field | description |
|---|---|
| `unique_id`, `model`, `resource_type`, `database`, `schema`, `alias` | identity |
| `materialized` | `view`, `table`, `incremental`, etc. |
| `path` | file path within the dbt project |
| `depends_on_nodes`, `depends_on_macros` | comma-joined upstream IDs |
| `raw_code` | the model's source SQL (with Jinja) |
| `compiled_code` | the fully resolved SQL used for lineage parsing |

### Column Lineage

The main event. One row per output column, **per stage** of each model's
query. See below for what "stage" means and how lineage is derived.

| field | description |
|---|---|
| `model`, `unique_id` | the model this column belongs to |
| `stage` | which part of the query defines this column - a CTE name, or `(final select)` |
| `is_final_output` | `True` if this stage's columns are the model's actual output columns |
| `column` | output column name (or `*` for an unexpanded star projection) |
| `expression` | the SQL expression that derives the column, **with query aliases resolved to real table/model names** (e.g. `stg_customers.customer_id` instead of `c.customer_id`) |
| `raw_expression` | the same expression exactly as written in the SQL, aliases included |
| `source_tables` | upstream table(s)/model(s) referenced by this expression |
| `source_columns` | upstream `table.column` references, comma-joined |
| `is_passthrough` | `True` if the column is a straight rename/copy with no transformation |
| `parse_note` | set when something couldn't be parsed or resolved (see Limitations) |

## How column lineage works

dbt models are commonly written as a chain of CTEs, e.g.:

```sql
with customers as (
    select * from {{ ref('stg_customers') }}
),
orders_summary as (
    select customer_id, sum(amount) as lifetime_spend
    from {{ ref('stg_orders') }}
    group by 1
),
final as (
    select c.customer_id, c.customer_name, o.lifetime_spend
    from customers c
    left join orders_summary o on c.customer_id = o.customer_id
)
select * from final
```

For each model, the script:

1. Parses `compiled_code` with `sqlglot`.
2. Treats **every named CTE as its own lineage stage** - so `customers`,
   `orders_summary`, and `final` above each get their own rows, letting you
   see multi-step transformations column by column instead of only the
   model's final output.
3. Detects the common `select * from final` pattern at the end of a model
   and marks that CTE's columns as the model's real output (`is_final_output
   = True`) instead of reporting a bare, unexpanded `*`.
4. For each output column's expression, resolves the query aliases used in
   its `FROM`/`JOIN` clauses back to real relation names:
   - A **transparent pass-through CTE** (`some_cte as (select * from
     {{ ref(...) }})`, no join, no aggregation) is resolved straight through
     to the real upstream model/source name - recursively, through chains of
     such CTEs.
   - A CTE that joins, aggregates, or otherwise transforms its input (e.g.
     `select order_id, sum(amount) ... group by order_id`) is **left as its
     own CTE name** rather than resolved further. Resolving through an
     aggregation would misattribute a computed column (like a `SUM`) to a
     raw upstream column that doesn't actually exist there.

The result: a column like `final.lifetime_spend` correctly traces back
through `orders_summary.lifetime_spend` (an aggregate) rather than being
incorrectly flattened all the way to a raw `stg_orders.lifetime_spend`
column that was never real.

## Limitations

- **Single-model scope.** Lineage is derived from each model's own compiled
  SQL. It does not recursively parse upstream models to trace a column all
  the way back through multiple models to its ultimate raw source - it
  resolves one model's CTEs plus its immediate `ref()`/`source()` upstreams.
- **`SELECT *` from a real table.** If a model selects `*` directly from a
  ref'd/sourced table (not a CTE) rather than naming columns, the script
  reports a single `*` row with a `parse_note` rather than expanding it,
  since the manifest doesn't always carry that table's full column list.
- **Complex/dialect-specific SQL.** Very unusual SQL (recursive CTEs, lateral
  joins, dialect-specific syntax `sqlglot` doesn't support) may fail to
  parse; affected models get a `parse_note` explaining the parse error
  instead of lineage rows.
- **Best-effort alias resolution.** Table/relation matching is done via
  name/schema/database string matching against the manifest, not live
  warehouse introspection - unusual quoting or naming conventions can
  occasionally leave a reference unresolved (falls back to the raw name
  found in the SQL).

## Troubleshooting

**"No nodes matched the given --model/--path filter(s)"** - the name/path
you passed doesn't match any node. Names are matched case-insensitively but
must be exact; paths accept glob patterns (`fnmatch` syntax) or plain
substrings.

**A model's lineage rows are all `parse_note` / empty** - either `sqlglot`
isn't installed, the manifest lacks `compiled_code` (re-run `dbt compile`),
or the SQL uses syntax `sqlglot` can't parse for your adapter's dialect.

**Excel file won't open** - re-download/re-copy the file; if the problem
persists, note the exact app and error message, since a truncated or
partial copy is the most common cause and is not a sign the workbook itself
is malformed.

## Supported adapters (for SQL dialect detection)

Snowflake, BigQuery, Redshift, Postgres, DuckDB, Databricks, Spark,
Trino, Presto, SQL Server. The adapter is read from
`manifest["metadata"]["adapter_type"]`; other adapters fall back to
`sqlglot`'s generic SQL dialect, which handles most standard syntax fine.
