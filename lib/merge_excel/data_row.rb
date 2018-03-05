module MergeExcel
  class DataRow
    attr_reader :array
    
    def initialize(array, extra_data_hash, book)
      if extra_data_hash
        extra_array = extra_data_hash.data.map do |e|
          case e.type
          when :filename
            book.filename
          when :cell_value
            book.cell_value_at(e.sheet_idx, e.row_idx, e.col_idx)
          else
            # error?
          end
        end
        case extra_data_hash.position
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
