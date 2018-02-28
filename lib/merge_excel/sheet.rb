module MergeExcel
  class Sheet
    attr_reader :original_idx, :idx, :name, :header, :data_rows
    def initialize(original_idx, idx, name, book)
      @original_idx = original_idx
      @idx = idx
      @name = name
      @header = nil
      @book = book
      @data_rows = []
    end

    def add_header(array, extra_data)
      @header = Header.new(array, extra_data)
    end

    def add_data_row(array, extra_data, book_filename, book)
      @data_rows << DataRow.new(array || [], extra_data, book_filename, book).array
    end
  end
end
