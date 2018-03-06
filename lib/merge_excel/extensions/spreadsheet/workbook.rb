module Spreadsheet
  class Workbook
    def add_worksheet_at(idx, name)
      ws = Worksheet.new(name: name)
      add_worksheet(ws)
    end

    def cell_value_at(sheet_idx, row_idx, col_idx)
      worksheet(sheet_idx).row(row_idx)[col_idx]
    end

    def count_worksheets
      worksheets.size
    end
  end
end
