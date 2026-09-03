import openpyxl

class XLSMHandler:
    def __init__(self, file_path):
        self.file_path = file_path
        # keep_vba=True is crucial for preserving the .xlsm macros when we save later
        self.workbook = openpyxl.load_workbook(file_path, keep_vba=True)
        self.data_sheet = self.workbook["Data"]
        
        # The user's dataset has its header on row 2 (index 1 if 0-indexed, but openpyxl is 1-indexed)
        # So row 2 is the header
        self.header_row_idx = 2
        
        self.headers = []
        for cell in self.data_sheet[self.header_row_idx]:
            self.headers.append(cell.value)
            
    def get_rows(self):
        """Yields each row as a dict, along with its original openpyxl row index for updating."""
        for row_idx in range(self.header_row_idx + 1, self.data_sheet.max_row + 1):
            row_data = {}
            for col_idx, cell in enumerate(self.data_sheet[row_idx], start=1):
                header_name = self.headers[col_idx - 1]
                row_data[header_name] = cell.value
                
            # Only yield rows that actually have an ID or some primary identifier
            if row_data.get("Platform") == "Trustpilot":
                yield row_idx, row_data

    def save(self, output_path):
        self.workbook.save(output_path)
