module MergeExcel
  class Header
    def initialize(array, extra_col_hash)

      if extra_col_hash
        extra_array = extra_col_hash.data.map{|e| e.label}
        case extra_col_hash.position
        when :beginning
          @array = extra_array + array
        else
          @array = array + extra_array
        end
      else
        @array = array
      end
    end

    def to_a
      @array.to_a
    end
  end
end
