import dotenv
import os
import sqlite3
import sys
from loguru import logger
import pandas as pd

dotenv.load_dotenv()

logger.remove(0)
logger.add(
    "./activity.log",
    format="{time}: {level} {message}",
    colorize=True,
    enqueue=True,
)
logger.add(
    sys.stderr,
    format="{time}: {level} {message}",
    colorize=True,
    enqueue=True,
)


def export_data(db_path: str, tables: list[str], xlsx_path: str) -> None:
    try:
        logger.info("Exporting data...")
        # Connect to SQLite database
        conn = sqlite3.connect(db_path)

        # Create Excel writer object
        with pd.ExcelWriter(xlsx_path) as writer:
            # Export each table to a different sheet
            for table in tables:
                query = f"SELECT * FROM {table}"
                df = pd.read_sql_query(query, conn)
                df.to_excel(writer, sheet_name=table, index=False)
                logger.info(f"Exported table {table} to sheet {table}")

        logger.info(f"Successfully exported all tables to {xlsx_path}")

    except Exception as e:
        logger.error(f"Error: {str(e)}")

    finally:
        # Close the connection
        conn.close()


# Usage
db_path = os.getenv("SQLITE_PATH", None)
tables = [x.strip() for x in os.getenv("TABLES", "").split(",") if x.strip()]
xlsx_path = os.getenv("OUTPUT_PATH", "output.xlsx")

if __name__ == "__main__":
    if not db_path or not os.path.exists(db_path):
        logger.error(f"Error: {db_path} does not exist.")
        sys.exit(1)
    if len(tables) == 0:
        logger.error("Error: No tables specified.")
        sys.exit(1)
    export_data(db_path, tables, xlsx_path)
