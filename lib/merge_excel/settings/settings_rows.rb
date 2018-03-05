# TO TRASH
module MergeExcel
  class SettingsRows
    def initialize
      @hash = {}
    end

    def insert(idx, h)
      if idx.is_a? Integer
        raise "Problem with '#{setting_name}' settings: sheet indexes must be >=0" if idx<0
      end
      @hash[idx] = row_class.new(h)
    end

    def get(sheet_idx) # a number>=0 or :any
      @hash.fetch(sheet_idx) do
        @hash.fetch(:any) do
          @hash[:any] = row_class.with_defaults
        end
      end
    end
  end
end
