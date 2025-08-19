import oracledb
import pandas as pd
import re


class OracleDB:
    def __init__(self, user, password, dsn):
        """Inicjalizacja połączenia z bazą Oracle"""
        self.conn = oracledb.connect(user=user, password=password, dsn=dsn)
        self.cursor = self.conn.cursor()

    def _table_exists(self, table_name, schema=None):
        """Sprawdza czy tabela istnieje"""
        if schema:
            self.cursor.execute("""
                SELECT table_name 
                FROM all_tables 
                WHERE owner = :sch AND table_name = :tbl
            """, [schema.upper(), table_name.upper()])
        else:
            self.cursor.execute("""
                SELECT table_name 
                FROM all_tables 
                WHERE table_name = :tbl
            """, [table_name.upper()])
        return self.cursor.fetchone() is not None


    def _get_table_columns(self, table_name, schema=None):
        """Pobiera wszystkie kolumny tabeli"""
        if schema:
            self.cursor.execute("""
                SELECT column_name 
                FROM all_tab_columns 
                WHERE owner = :sch AND table_name = :tbl
            """, [schema.upper(), table_name.upper()])
        else:
            self.cursor.execute("""
                SELECT column_name 
                FROM all_tab_columns 
                WHERE table_name = :tbl
            """, [table_name.upper()])
        return [row[0] for row in self.cursor.fetchall()]


    def _print_table_schema(self, table_name):
        """Wypisuje cały schemat tabeli"""
        self.cursor.execute("""
            SELECT column_name, data_type, data_length, nullable 
            FROM all_tab_columns 
            WHERE table_name = :tbl
            ORDER BY column_id
        """, [table_name.upper()])
        schema = self.cursor.fetchall()
        print(f"\nSchemat tabeli {table_name}:")
        for col in schema:
            print(f"{col[0]} - {col[1]}({col[2]}) - Nullable: {col[3]}")
        print("-" * 40)

    def execute_query(self, sql, save_to_xlsx=False, xlsx_filename="result.xlsx"):
        """Wykonuje zapytanie i opcjonalnie zapisuje do xlsx"""
        try:
            # jeśli SELECT z tabeli, sprawdzamy tabelę i kolumny
            # match = re.search(r"FROM\s+(\w+)", sql, re.IGNORECASE)
            match = re.search(r"FROM\s+([A-Z0-9_\.\"']+)", sql, re.IGNORECASE)

            if match:
                table_name = match.group(1).strip().upper()

                # jeśli jest w formacie SCHEMA.TABLE
                if "." in table_name:
                    schema, table = table_name.split(".")
                else:
                    schema = None
                    table = table_name

                if not self._table_exists(table, schema):
                    print(f"❌ Tabela {table_name} nie istnieje!")
                    return None
                else:
                    print(f"✅ Tabela {table_name} istnieje.")

                    # sprawdzamy kolumny
                    table_cols = self._get_table_columns(table_name)

                    # znajdź kolumny użyte w SELECT
                    select_match = re.search(r"SELECT\s+(.*?)\s+FROM", sql, re.IGNORECASE | re.DOTALL)
                    if select_match:
                        cols_in_query = [c.strip().upper() for c in select_match.group(1).split(",")]
                        if "*" not in cols_in_query:
                            missing = [c for c in cols_in_query if c not in table_cols]
                            if missing:
                                print(f"⚠️ Brakujące kolumny w tabeli {table_name}: {missing}")
                                self._print_table_schema(table_name)

            # wykonanie zapytania
            self.cursor.execute(sql)

            if self.cursor.description:  # SELECT
                columns = [col[0] for col in self.cursor.description]
                rows = self.cursor.fetchall()
                df = pd.DataFrame(rows, columns=columns)

                if save_to_xlsx:
                    df.to_excel(xlsx_filename, index=False)
                    print(f"💾 Wynik zapisany do pliku {xlsx_filename}")

                return df
            else:  # DML/DDL (INSERT, UPDATE, DELETE itp.)
                self.conn.commit()
                print("✅ Zapytanie wykonane pomyślnie.")
                return None

        except Exception as e:
            print(f"❌ Błąd podczas wykonywania zapytania: {e}")
            return None

    def close(self):
        """Zamyka połączenie"""
        self.cursor.close()
        self.conn.close()
