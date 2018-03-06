module MergeExcel
  module Settings
    class Parser
      attr_reader :input_dir, :output_file_path, :sheets_idxs, :data_rows, :extra_data_rows
      def initialize(hash)
        @input_dir         = hash.fetch(:input_dir)
        @output_file_path  = hash.fetch(:output_file_path)
        read_sheets_index    hash.fetch(:sheets_idxs){ :all }
        read_data_rows       hash.fetch(:data_rows){ {any: DataRow::DEFAULT_DATA_ROW_CONFIG } }
        read_extra_data_rows hash.fetch(:extra_data_rows){   {any: ExtraDataRow::DEFAULT_EXTRA_DATA_ROW_CONFIG } }
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

      def read_extra_data_rows(h)
        # {:any=>{:position=>:beginning, :data=>[{:type=>:filename, :heading_text=>"Filename"}}]}
        @extra_data_rows = ExtraDataRows.new
        h.each_pair do |k,v|
          @extra_data_rows.insert(k, v)
        end
      end
    end
  end
end
