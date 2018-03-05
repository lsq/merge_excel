# TO_TRASH
module MergeExcel
  module Settings
    class HeaderRow < SettingsRow
      def self.with_defaults
        new(DEFAULT_HEADER_ROW_CONFIG)
      end
      def setting_name
        "header_rows"
      end
    end
  end
end
