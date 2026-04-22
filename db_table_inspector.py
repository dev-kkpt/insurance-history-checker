import argparse
from datetime import datetime

import pandas as pd
import pyodbc


CONNECTION_STRING = (
    "DRIVER={SQL Server};"
    "SERVER=127.0.0.1,1436;"
    "DATABASE=DentWeb;"
    "UID=sa;"
    "PWD=Q3xzJiwpv2zC;"
    "TrustServerCertificate=yes;"
)


TABLE_SUMMARY_QUERY = """
SELECT
    s.name AS schema_name,
    t.name AS table_name,
    t.create_date,
    t.modify_date,
    COALESCE(SUM(CASE WHEN ps.index_id IN (0, 1) THEN ps.row_count ELSE 0 END), 0) AS row_count
FROM sys.tables AS t
INNER JOIN sys.schemas AS s
    ON s.schema_id = t.schema_id
LEFT JOIN sys.dm_db_partition_stats AS ps
    ON ps.object_id = t.object_id
GROUP BY
    s.name,
    t.name,
    t.create_date,
    t.modify_date
ORDER BY
    s.name,
    t.name;
"""


COLUMN_QUERY = """
SELECT
    TABLE_SCHEMA AS schema_name,
    TABLE_NAME AS table_name,
    ORDINAL_POSITION AS column_order,
    COLUMN_NAME AS column_name,
    DATA_TYPE AS data_type,
    CHARACTER_MAXIMUM_LENGTH AS max_length,
    IS_NULLABLE AS is_nullable
FROM INFORMATION_SCHEMA.COLUMNS
ORDER BY
    TABLE_SCHEMA,
    TABLE_NAME,
    ORDINAL_POSITION;
"""


def get_connection():
    return pyodbc.connect(CONNECTION_STRING)


def load_dataframe(cursor, query):
    cursor.execute(query)
    columns = [column[0] for column in cursor.description]
    return pd.DataFrame.from_records(cursor.fetchall(), columns=columns)


def print_table_summary(summary_df, columns_df):
    if summary_df.empty:
        print("DB에서 사용자 테이블을 찾지 못했습니다.")
        return

    for _, row in summary_df.iterrows():
        schema_name = row["schema_name"]
        table_name = row["table_name"]
        table_columns = columns_df[
            (columns_df["schema_name"] == schema_name)
            & (columns_df["table_name"] == table_name)
        ]
        print(f"\n[{schema_name}.{table_name}]")
        print(
            "row_count={row_count}, created={created}, modified={modified}".format(
                row_count=row["row_count"],
                created=row["create_date"],
                modified=row["modify_date"],
            )
        )
        for _, column in table_columns.iterrows():
            nullable = "NULL" if column["is_nullable"] == "YES" else "NOT NULL"
            length = (
                ""
                if pd.isna(column["max_length"])
                else f"({int(column['max_length'])})"
            )
            print(
                f"  - {column['column_order']:>2}. {column['column_name']} "
                f"{column['data_type']}{length} {nullable}"
            )


def fetch_sample_rows(cursor, schema_name, table_name, limit):
    query = f"SELECT TOP {int(limit)} * FROM [{schema_name}].[{table_name}]"
    cursor.execute(query)
    columns = [column[0] for column in cursor.description]
    return pd.DataFrame.from_records(cursor.fetchall(), columns=columns)


def export_to_excel(summary_df, columns_df, sample_tables, sample_limit):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"db_table_inspection_{timestamp}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="table_summary", index=False)
        columns_df.to_excel(writer, sheet_name="columns", index=False)

        if sample_limit > 0:
            with get_connection() as conn:
                cursor = conn.cursor()
                for schema_name, table_name in sample_tables:
                    sample_df = fetch_sample_rows(cursor, schema_name, table_name, sample_limit)
                    sheet_name = f"{schema_name}_{table_name}"[:31]
                    sample_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\nExcel 파일로 저장했습니다: {output_path}")


def parse_table_name(value):
    parts = value.split(".", 1)
    if len(parts) != 2:
        raise argparse.ArgumentTypeError(
            "샘플 대상 테이블은 schema.table 형식으로 입력해야 합니다."
        )
    return parts[0], parts[1]


def main():
    parser = argparse.ArgumentParser(
        description="DentWeb SQL Server의 테이블 구조와 데이터 현황을 점검합니다."
    )
    parser.add_argument(
        "--sample-limit",
        type=int,
        default=0,
        help="샘플 데이터를 보고 싶을 때 테이블당 조회할 행 수입니다. 기본값은 0입니다.",
    )
    parser.add_argument(
        "--sample-table",
        action="append",
        type=parse_table_name,
        default=[],
        help="샘플 데이터를 볼 테이블입니다. 예: dbo.TB_환자정보",
    )
    args = parser.parse_args()

    with get_connection() as conn:
        cursor = conn.cursor()
        summary_df = load_dataframe(cursor, TABLE_SUMMARY_QUERY)
        columns_df = load_dataframe(cursor, COLUMN_QUERY)

    print_table_summary(summary_df, columns_df)

    sample_tables = args.sample_table
    if args.sample_limit > 0 and not sample_tables:
        sample_tables = list(summary_df[["schema_name", "table_name"]].itertuples(index=False, name=None))[:5]
        print("\n샘플 테이블을 지정하지 않아 앞의 5개 테이블만 샘플로 저장합니다.")

    export_to_excel(summary_df, columns_df, sample_tables, args.sample_limit)


if __name__ == "__main__":
    main()
