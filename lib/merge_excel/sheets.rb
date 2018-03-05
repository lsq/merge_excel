module MergeExcel
  class Sheets
    attr_reader :array
    def initialize
      @array = []
    end

    def <<(element)
      @array << element
    end

    def each
      @array.each{|e| yield e}
    end

    def each_with_index
      @array.each_with_index{|e, i| yield(e,i)}
    end

    def get_sheet(original_idx: )
      @array.find{|s| s.source_sheet_idx==original_idx}
    end
  end
end
