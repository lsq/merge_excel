module MergeExcel
  module Settings
    class ExtraDataRows
      def initialize
        @hash = {}
      end

      def insert(sheet_idx, h)
        # h: {:any=>{:position=>:beginning, :data=>{:type=>:filename, :heading_text=>"Filename"}}}
        validate_sheet_idx sheet_idx
        @hash[sheet_idx] = ExtraDataRow.new(h)
      end

      def get(sheet_idx) # a number>=0 or :any
        validate_sheet_idx sheet_idx
        @hash.fetch(sheet_idx) do
          @hash.fetch(:any) do
            @hash[:any] = ExtraDataRow.with_defaults
          end
        end
      end

      private
      def validate_sheet_idx(sheet_idx)
        if sheet_idx.is_a?(Integer)
          raise "Problem with 'extra_data_rows' settings: sheet indexes must be >=0" if sheet_idx<0
        elsif sheet_idx!=:any
          raise "Problem with 'extra_data_rows' settings: sheet indexes must be an integer or :any symbol"
        end
      end
    end
  end
end
