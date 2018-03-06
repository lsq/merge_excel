module MergeExcel
  class Excel
    attr_reader :files

    def initialize(hash, selector_pattern="*.{xls,xlsx}")
      @settings = Settings::Parser.new(hash)
      @files   = Dir.glob(File.join(@settings.input_dir, selector_pattern))
    end

    def merge
      wbook = WBook.new(@settings)
      @files.each do |i_xls_filepath|
        puts i_xls_filepath
        wbook.import_data(i_xls_filepath)
      end
      wbook.export @settings.output_file_path
    end
  end
end
