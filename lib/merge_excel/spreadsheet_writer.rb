module MergeExcel
  class SpreadsheetWriter
    extend FormatDetector

    def self.create(filepath)
      extname = detect_format(filepath)
      wb = case extname
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
