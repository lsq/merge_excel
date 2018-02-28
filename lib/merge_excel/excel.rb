module MergeExcel
  class Excel
    attr_reader :files
    def initialize(hash, selector_pattern="*.{xls,xlsx}")
      @h     = hash
      @files = Dir.glob(File.join(@h[:input_dir], selector_pattern))
      check_options
    end

    def merge
      wbook = WBook.new(@files.first, @h)
      @files.each do |i_xls_filepath|
        puts i_xls_filepath
        wbook.import_data(i_xls_filepath)
      end
      output_file_path = @h[:output_file_path]
      wbook.export output_file_path
    end


    private

    def check_options
      @h = {
        sheets_idxs: :all,
        header_rows: {
          any: {
            row_idx: 0,
            from_col_idx: 0,
            to_col_idx: :last
          }
        },
        data_rows: {
          any: {
            first_row_idx: 1,
            from_col_idx: 0,
            to_col_idx: :last
          }
        },
        extra_data: {}
      }.merge(@h)
    end
  end
end
