module RubyXL
  class Worksheet
    def row_values(row_idx, from_col_idx=0, to_col_idx=-1)
      values = sheet_data[row_idx] && sheet_data[row_idx].cells[from_col_idx..to_col_idx].map{|c| c && c.value}
      return nil if values.nil? || values.empty? || values.all?(&:nil?)
      values
    end

    def add_row_at(row_idx, cells_array)
      cells_array.to_a.each_with_index do |cell, col_idx|
        if cell.is_a? DateTime
          c = add_cell(row_idx, col_idx)
          c.set_number_format('yyyy-mm-dd')
          c.change_contents(cell)
        else
          add_cell(row_idx, col_idx, cell)
        end
      end
    end
  end
end
