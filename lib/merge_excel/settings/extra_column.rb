module MergeExcel
  module Settings
    class ExtraColumn
      DEFAULT_EXTRA_COLUMN_CONFIG = {}

      class FilenameExtraColumn
        TYPE = :filename
        attr_reader :label, :type
        def initialize(label)
          @label = label
          @type  = :filename
        end
      end

      class CellValueExtraColumn
        TYPE = :cell_value
        attr_reader :label, :type, :sheet_idx, :row_idx, :col_idx
        def initialize(label, coordinates_hash)
          @label     = label
          @sheet_idx = coordinates_hash.fetch(:sheet_idx)
          @row_idx   = coordinates_hash.fetch(:row_idx)
          @col_idx   = coordinates_hash.fetch(:col_idx)
          @type      = :cell_value
        end
      end


      attr_reader :position, :data
      def initialize(hash)
        # hash: {:position=>:beginning, :data=>[{:type=>:filename, :heading_text=>"Filename"}]}
        @position = hash.fetch(:position){ :beginning }
        @data = []
        read_data_array hash.fetch(:data){ [] }
        validate_position
      end

      def read_data_array(array)
        array.each do |h|
          if h[:type]==:filename
            @data << FilenameExtraColumn.new(h.fetch(:label){"Filename"})
          elsif h[:type]==:cell_value
            @data << CellValueExtraColumn.new(h.fetch(:label){"Field"}, h)
          end
        end
      end


      def self.with_defaults
        new(DEFAULT_EXTRA_COLUMN_CONFIG)
      end

      private
      def validate_position
        raise "Problem with 'extra_column' settings: 'position' must be :beginning or :end" unless [:beginning, :end].include?(@position)
      end
    end
  end
end
