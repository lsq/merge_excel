module MergeExcel
  class Excel
    attr_reader :files

    def initialize(input_dir, options={})
      @settings = Settings::Parser.new(options)
      @files    = Dir.glob(File.join(input_dir, @settings.selector_pattern))
    end

    def merge(output_file_path)
      wbook = WBook.new(@settings)
      @files.each do |i_xls_filepath|
        puts i_xls_filepath
        wbook.import_data(i_xls_filepath)
      end
      wbook.export output_file_path
    end
  end
end
