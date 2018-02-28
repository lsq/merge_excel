module MergeExcel
  class DataRow
    attr_reader :array
    def initialize(array, extra_data_hash, book_filename, book)
      if extra_data_hash
        extra_array = extra_data_hash[:data].map do |e|
          case e[:data]
          when :filename
            book_filename
          else
            cell_value_at(e[:data][:sheet_idx], e[:data][:row_idx], e[:data][:col_idx])
          end
        end
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
