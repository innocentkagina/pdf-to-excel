import os
import logging
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")


def export_excels_to_postgres(
    folder_path,
    POSTGRES_HOST=None,
    POSTGRES_PORT=None,
    POSTGRES_DB=None,
    POSTGRES_USER=None,
    POSTGRES_PASSWORD=None,
    schema="pdf_to_excel_data",
    if_exists="replace"
):
    """
    Export all Excel files in `folder_path` to PostgreSQL tables within the given schema.
    Adds 'serialno' and 'date_imported' columns.

    Parameters:
        folder_path (str): Path to directory containing .xlsx files.
        POSTGRES_HOST (str): Database host. Defaults to env POSTGRES_HOST or 'postgres'.
        POSTGRES_PORT (int): Database port. Defaults to env POSTGRES_PORT or 5432.
        POSTGRES_DB (str): Database name. Defaults to env POSTGRES_DB or 'airflow'.
        POSTGRES_USER (str): Database user. Defaults to env POSTGRES_USER or 'airflow'.
        POSTGRES_PASSWORD (str): Database password. Defaults to env POSTGRES_PASSWORD or 'airflow'.
        schema (str): PostgreSQL schema to write tables into. Default "pdf_to_excel_data".
        if_exists (str): Behavior when table exists - "replace" (default), "append", or "fail".
    """

    # Load configuration from env if not provided
    POSTGRES_HOST = POSTGRES_HOST or os.getenv("POSTGRES_HOST", "localhost")
    POSTGRES_PORT_env = POSTGRES_PORT or os.getenv("POSTGRES_PORT", "5432")
    try:
        POSTGRES_PORT = int(POSTGRES_PORT_env)
    except (TypeError, ValueError):
        logging.warning(f"Invalid POSTGRES_PORT '{POSTGRES_PORT_env}', defaulting to 5432")
        POSTGRES_PORT = 5432
    POSTGRES_DB = POSTGRES_DB or os.getenv("POSTGRES_DB")
    POSTGRES_USER = POSTGRES_USER or os.getenv("POSTGRES_USER")
    POSTGRES_PASSWORD = POSTGRES_PASSWORD or os.getenv("POSTGRES_PASSWORD")

    # Build DB connection string
    db_url = f"postgresql+psycopg2://{POSTGRES_USER}:{POSTGRES_PASSWORD}@{POSTGRES_HOST}:{POSTGRES_PORT}/{POSTGRES_DB}"

    try:
        engine = create_engine(db_url, future=True)
        with engine.connect() as conn:
            conn.execute(text(f'CREATE SCHEMA IF NOT EXISTS "{schema}"'))
            conn.commit()
    except SQLAlchemyError as e:
        logging.error(f"Error connecting to DB or creating schema: {e}")
        raise

    # Process each Excel file
    for fname in os.listdir(folder_path):
        if fname.lower().endswith(".xlsx"):
            table_name = os.path.splitext(fname)[0].lower()
            file_path = os.path.join(folder_path, fname)
            try:
                df = pd.read_excel(file_path, engine="openpyxl")
                if df.empty:
                    logging.warning(f"{fname} is empty, skipping.")
                    continue

                # Clean columns: lowercase and underscores instead of spaces
                df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

                # Add serialno and date_imported columns
                df["serialno"] = range(1, len(df) + 1)
                df["date_imported"] = datetime.now()

                # with engine.begin() as conn:
                df.to_sql(
                    table_name,
                    con=engine,
                    schema=schema,
                    if_exists=if_exists,
                    index=False,
                    method="multi",
                )
                logging.info(f"Exported {fname} to {schema}.{table_name}")
            except Exception as e:
                logging.error(f"Failed to export {fname}: {e}")
