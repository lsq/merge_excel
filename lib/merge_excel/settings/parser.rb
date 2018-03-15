module MergeExcel
  module Settings
    class Parser
      attr_reader :selector_pattern, :sheets_idxs, :data_rows, :extra_columns
      def initialize(hash)
        @selector_pattern  = hash.fetch(:selector_pattern){ "*.{xls,xlsx}" }
        read_sheets_index    hash.fetch(:sheets_idxs){ :all }
        read_data_rows       hash.fetch(:data_rows){ {any: DataRow::DEFAULT_DATA_ROW_CONFIG } }

        if hash.has_key?(:extra_data_rows) # legacy setting
          read_extra_columns   hash.fetch(:extra_data_rows){ {any: ExtraColumn::DEFAULT_EXTRA_COLUMN_CONFIG } }
        else
          read_extra_columns   hash.fetch(:extra_columns){ {any: ExtraColumn::DEFAULT_EXTRA_COLUMN_CONFIG } }
        end
      end

      def all_sheets?
        @sheets_idxs == :all
      end

      def read_sheet?(sheet_idx)
        all_sheets? || @sheets_idxs.include?(sheet_idx)
      end

      private

      def read_sheets_index(obj)
        if obj==:all || obj==[:all]
          @sheets_idxs = :all
        else
          obj = [obj] if !(obj.is_a? Array)
          @sheets_idxs = obj.select{|e| e.is_a? Integer}.sort
        end
      end

      def read_data_rows(h)
        @data_rows = DataRows.new
        h.each_pair do |k,v|
          @data_rows.insert(k, v)
        end
      end

      def read_extra_columns(h)
        # {:any=>{:position=>:beginning, :data=>[{:type=>:filename, :heading_text=>"Filename"}}]}
        @extra_columns = ExtraColumns.new
        h.each_pair do |k,v|
          @extra_columns.insert(k, v)
        end
      end
    end
  end
end
