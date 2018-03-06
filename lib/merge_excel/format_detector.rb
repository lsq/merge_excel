module MergeExcel
  module FormatDetector
    def detect_format(filepath)
      case File.extname(filepath)
      when ".xls"
        :xls
      when ".xlsx"
        :xlsx
      else
        raise "Invalid format"
      end
    end
  end
end
