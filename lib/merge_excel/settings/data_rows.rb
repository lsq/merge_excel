module MergeExcel
  module Settings
    class DataRows
      def initialize
        @hash = {}
      end

      def insert(sheet_idx, h)
        validate_sheet_idx sheet_idx
        @hash[sheet_idx] = DataRow.new(h)
      end

      def get(sheet_idx) # a number>=0 or :any
        validate_sheet_idx sheet_idx
        @hash.fetch(sheet_idx) do
          @hash.fetch(:any) do
            @hash[:any] = DataRow.with_defaults
          end
        end
      end

      private
      def validate_sheet_idx(sheet_idx)
        if sheet_idx.is_a?(Integer)
          raise "Problem with 'data_rows' settings: sheet indexes must be >=0" if sheet_idx<0
        elsif sheet_idx!=:any
          raise "Problem with 'data_rows' settings: sheet indexes must be an integer or :any symbol"
        end
      end
    end
  end
end
