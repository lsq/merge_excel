module RubyXL
  class Workbook
    def add_worksheet_at(idx, name)
      sheet = Worksheet.new(sheet_name: name, workbook: self)
      worksheets[idx] = sheet
      sheet
    end

    def cell_value_at(sheet_idx, row_idx, col_idx)
      c = worksheets[sheet_idx][row_idx][col_idx]
      c && c.value
    end
  end
end
