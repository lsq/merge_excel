# TO TRASH
module MergeExcel
  module Settings
    class SettingsRow
      FIRST_COLUMN = 0
      LAST_COLUMN = -1
      DEFAULT_HEADER_ROW_CONFIG = {row_idx: 0, from_col_idx: FIRST_COLUMN, to_col_idx: LAST_COLUMN}
      DEFAULT_DATA_ROW_CONFIG   = {row_idx: 1, from_col_idx: FIRST_COLUMN, to_col_idx: LAST_COLUMN}

      attr_reader :row_idx, :from_col_idx, :to_col_idx
      def initialize(hash)
        @row_idx = hash.fetch(:row_idx){ 0 }
        @from_col_idx = hash.fetch(:from_col_idx){ FIRST_COLUMN }
        @to_col_idx = hash.fetch(:to_col_idx){ LAST_COLUMN }
        validate_row
        validate_columns
      end

      private
      def validate_row
        raise "Problem with '#{setting_name}' settings: row_idx must be >=0" if @row_idx<0
      end
      def validate_columns
        raise "Problem with '#{setting_name}' settings: from_col_idx must be >=0" if @from_col_idx<0
        raise "Problem with '#{setting_name}' settings: from_col_idx must be >=0" if @to_col_idx>=0 && @from_col_idx>@to_col_idx
      end
    end
  end
end
