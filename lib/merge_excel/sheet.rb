module MergeExcel
  class Sheet
    attr_reader :source_sheet, :source_sheet_idx, :sheet_idx, :name, :book, :header, :data_rows
    def initialize(source_sheet, source_sheet_idx)
      @source_sheet     = source_sheet
      @source_sheet_idx = source_sheet_idx
      @name             = @source_sheet.sheet_name
      @book             = @source_sheet.workbook

      @header = nil
      @data_rows = []
    end

    def add_header(array, extra_data)
      @header = Header.new(array, extra_data)
    end

    def add_data_row(array, extra_data, book)
      @data_rows << DataRow.new(array || [], extra_data, book).array
    end
  end
end
