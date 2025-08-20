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
