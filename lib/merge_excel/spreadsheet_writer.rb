module MergeExcel
  class SpreadsheetWriter
    def self.create(filepath, format)
      wb = case format
      when :xls
        Spreadsheet::Workbook.new

      when :xlsx
        RubyXL::Workbook.new
      else
        raise "Invalid format"
      end

      yield(wb)
      wb.write filepath
    end
  end
end
