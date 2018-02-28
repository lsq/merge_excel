module MergeExcel
  class SpreadsheetParser
    def self.open(filepath, format)
      case format
      when :xls
        Spreadsheet.open(filepath)
      when :xlsx
        RubyXL::Parser.parse(filepath)
      else
        raise "Invalid format"
      end
    end
  end
end
