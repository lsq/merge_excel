# TO_TRASH
module MergeExcel
  module Settings
    class HeaderRows < SettingsRows
      def setting_name
        "header_rows"
      end
      def row_class
        HeaderRow
      end
    end
  end
end
