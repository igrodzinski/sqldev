import pandas as pd
import re
import json
from openpyxl import load_workbook

class ExcelMappingProcessor:
    def __init__(self, filepath):
        self.filepath = filepath

    def _split_table_names(self, table_hd_value):
        if pd.isna(table_hd_value):
            return []
        tables = [t.strip() for t in str(table_hd_value).split(",") if t.strip()]
        result = []
        for t in tables:
            match = re.search(r"\((.*?)\)", t)
            comment = match.group(1).strip() if match else None
            clean_name = re.sub(r"\(.*?\)", "", t).strip()
            result.append((clean_name, comment))
        return result

    def process_and_save(self, output_sheet="processed"):
        df = pd.read_excel(self.filepath)
        new_rows = []
        for _, row in df.iterrows():
            tables = self._split_table_names(row["table_hd"])
            if not tables:
                new_row = row.to_dict()
                new_row["table_hd_comment"] = None
                new_rows.append(new_row)
            else:
                for table_name, comment in tables:
                    new_row = row.to_dict()
                    new_row["table_hd"] = table_name
                    new_row["table_hd_comment"] = comment
                    new_rows.append(new_row)
        df_processed = pd.DataFrame(new_rows)

        book = load_workbook(self.filepath)
        with pd.ExcelWriter(self.filepath, engine="openpyxl", mode="a") as writer:
            writer.book = book
            if output_sheet in writer.book.sheetnames:
                del writer.book[output_sheet]
            df_processed.to_excel(writer, sheet_name=output_sheet, index=False)
        print(f"✅ Utworzono nowy arkusz '{output_sheet}' w pliku {self.filepath}")
        return df_processed

    def rename_columns(self, df, mapping_json):
        """
        Zmienia nazwy kolumn w DataFrame zgodnie z mapowaniem.
        mapping_json: dict lub ścieżka do pliku JSON {"stara_nazwa": "nowa_nazwa"}
        """
        if isinstance(mapping_json, str):
            with open(mapping_json, "r", encoding="utf-8") as f:
                mapping = json.load(f)
        elif isinstance(mapping_json, dict):
            mapping = mapping_json
        else:
            raise ValueError("mapping_json musi być dict lub ścieżką do pliku JSON")

        df_renamed = df.rename(columns=mapping)
        return df_renamed

    def process_to_df(self):
        df = pd.read_excel(self.filepath)
        new_rows = []
        for _, row in df.iterrows():
            tables = self._split_table_names(row["table_hd"])
            if not tables:
                new_row = row.to_dict()
                new_row["table_hd_comment"] = None
                new_rows.append(new_row)
            else:
                for table_name, comment in tables:
                    new_row = row.to_dict()
                    new_row["table_hd"] = table_name
                    new_row["table_hd_comment"] = comment
                    new_rows.append(new_row)
        return pd.DataFrame(new_rows)


class TableIdentifierProcessor:
    def __init__(self, excel_path, sheet_name="processed_mapping"):
        self.excel_path = excel_path
        self.sheet_name = sheet_name

    def apply_identifications(self, json_path="tables_identificated.json", output_sheet="processed_identified"):
        """
        Wczytuje plik JSON i uzupełnia kolumnę table_identified na podstawie mapowania
        """
        # wczytaj dane
        df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name)

        # wczytaj mapowanie
        with open(json_path, "r", encoding="utf-8") as f:
            mapping = json.load(f)

        # dodaj kolumnę jeśli nie istnieje
        if "table_identified" not in df.columns:
            df["table_identified"] = None

        # uzupełnij kolumnę na podstawie mapowania
        df["table_identified"] = df.apply(
            lambda row: mapping.get(row["table_hd"], row["table_identified"]),
            axis=1
        )

        # zapisz do nowego arkusza
        with pd.ExcelWriter(self.excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=output_sheet, index=False)

        print(f"✅ Zaktualizowano kolumnę 'table_identified' i zapisano do arkusza '{output_sheet}'")
        return df

    def generate_queries(self, output_queries_noname="queries_noname.json", output_queries_def="queries_def3000.json"):
        """
        Tworzy listę unikalnych tabel bez identyfikacji i zapisuje SQL w plikach JSON
        """
        # wczytaj dane z excela (zakładamy, że apply_identifications już zrobiło nowy arkusz)
        df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name)

        # wybierz unikalne nazwy z table_hd, gdzie table_identified jest puste
        if "table_identified" in df.columns:
            df_missing = df[df["table_identified"].isna()]
        else:
            df_missing = df

        tables_unique = df_missing["table_hd"].dropna().unique().tolist()

        if not tables_unique:
            print("ℹ️ Brak tabel do przetworzenia (wszystkie mają table_identified).")
            return

        # zbuduj string 'T1','T2',...
        tables_unique_str = ",".join([f"'{t}'" for t in tables_unique])

        # SQL-e
        sql_noname = f"""
SELECT owner, table_name
FROM all_tables
WHERE table_name in ({tables_unique_str})
"""
        sql_def = f"""
SELECT table_name
FROM all_tables
WHERE owner = 'DEPOZ'
  AND table_name in ({tables_unique_str})
"""

        # zapis do JSON
        with open(output_queries_noname, "w", encoding="utf-8") as f:
            json.dump({",".join(tables_unique): sql_noname.strip()}, f, indent=4, ensure_ascii=False)

        with open(output_queries_def, "w", encoding="utf-8") as f:
            json.dump({",".join(tables_unique): sql_def.strip()}, f, indent=4, ensure_ascii=False)

        print(f"✅ Zapisano pliki {output_queries_noname} i {output_queries_def}")
        return tables_unique
