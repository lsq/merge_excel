module Spreadsheet
  class Worksheet
    def sheet_name
      name
    end

    def row_values(row_idx, from_col_idx=0, to_col_idx=-1)
      values = row(row_idx) && row(row_idx).to_a[from_col_idx..to_col_idx]
      return nil if values.nil? || values.empty? || values.all?(&:nil?)
      values
    end

    def add_row_at(row_idx, cells_array)
      row(row_idx).concat cells_array.to_a
    end
  end
end
