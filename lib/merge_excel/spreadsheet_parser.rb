module MergeExcel
  class SpreadsheetParser
    extend FormatDetector

    def self.open(filepath)
      extname = detect_format(filepath)
      case extname
      when :xls
        book = Spreadsheet.open(filepath)
      when :xlsx
        book = RubyXL::Parser.parse(filepath)
      end
      book.instance_variable_set(:@extname, extname)
      book.instance_variable_set(:@filename, File.basename(filepath))

      def book.xls?
        @extname==:xls ? true : false
      end
      def book.xlsx?
        !book.xls?
      end
      def book.filename
        @filename
      end
      book
    end
  end
end
