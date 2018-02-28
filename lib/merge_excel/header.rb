module MergeExcel
  class Header
    def initialize(array, extra_data_hash)
      if extra_data_hash
        extra_array = extra_data_hash[:data].map{|e| e[:heading_text]}
        case extra_data_hash[:position]
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
