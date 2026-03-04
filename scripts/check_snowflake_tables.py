#!/usr/bin/env python3
"""
check_snowflake_tables.py

Reads an Excel file containing report names and table names, connects to
Snowflake, and checks whether each table exists in the given database/schema.

Expected Excel format (one sheet, any name):
    Column A (or named "report_name")  – name of the report
    Column B (or named "table_name")   – fully-qualified or bare table name
                                         (DATABASE.SCHEMA.TABLE  or  TABLE)

Usage:
    python check_snowflake_tables.py \
        --input  reports.xlsx \
        --output results.xlsx \
        --account <account_identifier> \
        --user    <username> \
        --password <password> \
        --warehouse <warehouse> \
        --database  <database> \
        --schema    <schema>

Optional:
    --sheet         Sheet name or 0-based index (default: first sheet)
    --report-col    Column name for report names  (default: first column)
    --table-col     Column name for table names   (default: second column)
    --role          Snowflake role to use
"""

import argparse
import getpass
import os
import sys

import pandas as pd
import snowflake.connector


# ---------------------------------------------------------------------------
# Argument parsing
# ---------------------------------------------------------------------------

def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Bulk-check table existence in Snowflake from an Excel file."
    )
    parser.add_argument("--input", required=True,
                        help="Path to the input Excel file (.xlsx / .xls)")
    parser.add_argument("--output", default="results.xlsx",
                        help="Path for the output Excel file (default: results.xlsx)")
    parser.add_argument("--sheet", default=0,
                        help="Sheet name or 0-based index (default: 0)")
    parser.add_argument("--report-col", default=None,
                        help="Column header for report names (default: first column)")
    parser.add_argument("--table-col", default=None,
                        help="Column header for table names (default: second column)")

    # Snowflake connection
    parser.add_argument("--account", required=True,
                        help="Snowflake account identifier (e.g. xy12345.us-east-1)")
    parser.add_argument("--user", required=True,
                        help="Snowflake username")
    parser.add_argument("--password", default=None,
                        help="Snowflake password. If omitted, the "
                             "SNOWFLAKE_PASSWORD environment variable is used; "
                             "if that is also unset, you will be prompted.")
    parser.add_argument("--warehouse", required=True,
                        help="Snowflake warehouse to use")
    parser.add_argument("--database", required=True,
                        help="Default Snowflake database to search in")
    parser.add_argument("--schema", required=True,
                        help="Default Snowflake schema to search in")
    parser.add_argument("--role", default=None,
                        help="Snowflake role (optional)")

    args = parser.parse_args(argv)
    # Convert --sheet to int if it looks like a number
    try:
        args.sheet = int(args.sheet)
    except ValueError:
        pass  # keep as string (sheet name)

    return args


# ---------------------------------------------------------------------------
# Excel reading
# ---------------------------------------------------------------------------

def load_excel(path, sheet, report_col, table_col):
    """Return a DataFrame with columns 'report_name' and 'table_name'."""
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df.columns = df.columns.str.strip()

    cols = list(df.columns)
    if len(cols) < 2:
        sys.exit("ERROR: Excel sheet must have at least two columns.")

    # Resolve report column
    if report_col:
        if report_col not in cols:
            sys.exit(f"ERROR: report column '{report_col}' not found. "
                     f"Available columns: {cols}")
        r_col = report_col
    else:
        r_col = cols[0]

    # Resolve table column
    if table_col:
        if table_col not in cols:
            sys.exit(f"ERROR: table column '{table_col}' not found. "
                     f"Available columns: {cols}")
        t_col = table_col
    else:
        t_col = cols[1]

    result = df[[r_col, t_col]].copy()
    result.columns = ["report_name", "table_name"]
    result.dropna(subset=["table_name"], inplace=True)
    result["report_name"] = result["report_name"].fillna("")
    result["table_name"] = result["table_name"].str.strip()
    return result


# ---------------------------------------------------------------------------
# Snowflake connection
# ---------------------------------------------------------------------------

def connect_snowflake(args):
    password = (
        args.password
        or os.environ.get("SNOWFLAKE_PASSWORD")
        or getpass.getpass("Snowflake password: ")
    )
    connect_kwargs = dict(
        account=args.account,
        user=args.user,
        password=password,
        warehouse=args.warehouse,
        database=args.database,
        schema=args.schema,
    )
    if args.role:
        connect_kwargs["role"] = args.role
    return snowflake.connector.connect(**connect_kwargs)


# ---------------------------------------------------------------------------
# Table existence check
# ---------------------------------------------------------------------------

def parse_table_name(table_name, default_database, default_schema):
    """
    Split a table reference into (database, schema, table).

    Accepts:
        TABLE
        SCHEMA.TABLE
        DATABASE.SCHEMA.TABLE
    """
    parts = [p.strip().upper() for p in table_name.split(".")]
    if len(parts) == 1:
        return default_database.upper(), default_schema.upper(), parts[0]
    if len(parts) == 2:
        return default_database.upper(), parts[0], parts[1]
    if len(parts) == 3:
        return parts[0], parts[1], parts[2]
    # More than 3 dots – treat the whole string as the table name
    return default_database.upper(), default_schema.upper(), table_name.upper()


def check_tables(cursor, df, default_database, default_schema):
    """
    For every row in df, check whether the table exists in Snowflake's
    INFORMATION_SCHEMA.TABLES.

    Returns a list of dicts with the result for each row.
    """
    results = []
    for _, row in df.iterrows():
        report = row["report_name"]
        raw_table = row["table_name"]

        db, schema, tbl = parse_table_name(raw_table, default_database, default_schema)

        # Quote the database identifier to prevent SQL injection.
        # Snowflake identifier quoting: double any embedded double-quotes, then wrap.
        quoted_db = '"' + db.replace('"', '""') + '"'
        query = (
            "SELECT COUNT(*) FROM {db}.INFORMATION_SCHEMA.TABLES "
            "WHERE TABLE_SCHEMA = %s AND TABLE_NAME = %s"
        ).format(db=quoted_db)

        try:
            cursor.execute(query, (schema, tbl))
            count = cursor.fetchone()[0]
            exists = count > 0
            error = None
        except Exception as exc:  # pragma: no cover – network errors
            exists = False
            error = str(exc)

        results.append({
            "report_name": report,
            "table_name": raw_table,
            "database": db,
            "schema": schema,
            "table": tbl,
            "exists_in_snowflake": "YES" if exists else "NO",
            "error": error or "",
        })

        status = "FOUND" if exists else "NOT FOUND"
        print(f"[{status}]  {report!r:30s}  {raw_table}")

    return results


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(argv=None):
    args = parse_args(argv)

    print(f"Loading Excel file: {args.input}")
    df = load_excel(args.input, args.sheet, args.report_col, args.table_col)
    print(f"  {len(df)} rows found.\n")

    print("Connecting to Snowflake…")
    con = connect_snowflake(args)
    cursor = con.cursor()
    print("  Connected.\n")

    try:
        results = check_tables(cursor, df, args.database, args.schema)
    finally:
        cursor.close()
        con.close()

    out_df = pd.DataFrame(results)
    out_df.to_excel(args.output, index=False)

    found = sum(1 for r in results if r["exists_in_snowflake"] == "YES")
    print(f"\nDone. {found}/{len(results)} tables exist in Snowflake.")
    print(f"Results saved to: {args.output}")


if __name__ == "__main__":
    main()
