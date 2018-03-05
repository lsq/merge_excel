module MergeExcel
  module FormatDetector
    def detect_format(filepath)
      case File.extname(filepath)
      when ".xls"
        :xls
      when ".xlsx"
        :xlsx
      else
        raise "Invalid format"
      end
    end
  end
end

# module
#   module ClassMethods
#
#   end
#
#   module InstanceMethods
#
#   end
#
#   def self.included(receiver)
#     receiver.extend         ClassMethods
#     receiver.send :include, InstanceMethods
#   end
# end
