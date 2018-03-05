module MergeExcel
  module Settings
    class DataRow
      FIRST_COLUMN = 0
      LAST_COLUMN = -1
      DEFAULT_DATA_ROW_CONFIG = {
        header_row: 0,
        data_starting_row: 1,
        data_first_column: FIRST_COLUMN,
        data_last_column: LAST_COLUMN
      }

      attr_reader :header_row, :data_starting_row, :data_first_column, :data_last_column
      def initialize(hash)
        @header_row        = hash.fetch(:header_row){ 0 }
        @data_starting_row = hash.fetch(:data_starting_row){ 1 }
        @data_first_column = hash.fetch(:data_first_column){ FIRST_COLUMN }
        @data_last_column  = hash.fetch(:data_last_column){ LAST_COLUMN }
        @data_last_column  = LAST_COLUMN if @data_last_column==:last
        validate_rows
        validate_columns
      end

      def import_header?
        @header_row!=:missing
      end

      # def self.with_defaults
      #   new(DEFAULT_HEADER_ROW_CONFIG)
      # end

      private
      def validate_rows
        # header_row
        if (@header_row.is_a?(Symbol) && @header_row!=:missing) || (@header_row.is_a?(Integer) && @header_row<0)
          raise "Problem with 'data_rows' settings: header_row must be be an integer >=0 or :missing symbol" if @header_row<0
        end
        # data_starting_row
        if @data_starting_row<0
          raise "Problem with 'data_rows' settings: data_starting_row must be be an integer >=0" if @data_starting_row<0
        end
        # header_row vs data_starting_row
        if @header_row!=:missing && (@header_row >= @data_starting_row)
          raise "Problem with 'data_rows' settings: header_row must be < data_starting_row" if @header_row<0
        end
      end
      def validate_columns
        raise "Problem with 'data_rows' settings: data_first_column must be >=0" if @data_first_column<0
        raise "Problem with 'data_rows' settings: data_last_column must be >=0" if @data_last_column>=0 && @data_first_column>@data_last_column
      end
    end
  end
end
