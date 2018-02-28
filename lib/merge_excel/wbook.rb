module MergeExcel
  class WBook
    attr_reader :sheets, :hash, :format, :filepath, :filename

    def initialize(filepath, hash)
      @sheets   = []
      @hash     = hash
      @format   = detect_format(filepath) # :xls or :xlsx
      @filepath = filepath
      @filename = File.basename(filepath)

      book = SpreadsheetParser.open(filepath, @format)
      book.worksheets.each_with_index do |sheet, idx|
        i = 0
        if @hash[:sheets_idxs]==:all || @hash[:sheets_idxs].include?(idx)
          raise InvalidOptions if !@hash[:header_rows].keys.include?(idx)
          s = @hash[:header_rows][idx] # settings
          sh = Sheet.new(idx, i, sheet.sheet_name, self)
          arr = sheet.row_values(s[:row_idx], s[:from_col_idx], s[:to_col_idx])
          sh.add_header(arr, @hash[:extra_data][idx])
          @sheets << sh
          i+=1
        end
      end
    end


    def sheet(original_idx: )
      @sheets.find{|s| s.original_idx==original_idx}
    end


    def import_data(filepath)
      book = SpreadsheetParser.open(filepath, @format)
      filename = File.basename(filepath)
      @hash[:sheets_idxs].each do |idx|
        s = sheet(original_idx: idx)
        sheet_to_read = book.worksheets[idx]
        if row_settings = @hash[:data_rows][idx] # is an hash
          row_idx = row_settings[:first_row_idx]
          loop do
            break unless cells = sheet_to_read.row_values(row_idx, row_settings[:from_col_idx], row_settings[:to_col_idx])
            s.add_data_row(cells, @hash[:extra_data][idx], filename, book)
            row_idx+=1
          end
        end
      end
    end


    def export(filepath)
      print "Exporting data..."
      SpreadsheetWriter.create(filepath, @format) do |wb|
        @sheets.each_with_index do |sheet, sheet_idx|
          s = wb.add_worksheet_at(sheet_idx, sheet.name)
          s.add_row_at(0, sheet.header)
          sheet.data_rows.each_with_index do |data_row, row_idx|
            s.add_row_at(row_idx+1, data_row)
          end
        end
      end

      # if xls?
        # merged_book = Spreadsheet::Workbook.new
        # @sheets.each do |sheet|
        #   s = merged_book.create_worksheet name: sheet.name
        #   s.row(0).concat sheet.header.to_a
        #   sheet.data_rows.each_with_index do |data_row, i|
        #     s.row(i+1).concat data_row.to_a
        #   end
        # end
        # merged_book.write filename
      #
      # else
      #   merged_book = RubyXL::Workbook.new
      #   @sheets.each_with_index do |sheet, i|
      #     if i==0
      #       s = merged_book.worksheets[0]
      #       s.sheet_name = sheet.name
      #     else
      #       s = merged_book.add_worksheet sheet.name
      #     end
      #
      #     sheet.header.to_a.each_with_index do |header_cell, idx|
      #       s.add_cell(0, idx, header_cell)
      #     end
      #     sheet.data_rows.each_with_index do |data_row, row_idx|
      #       data_row.to_a.each_with_index do |data_cell, col_idx|
      #         # puts "#{data_cell} -> #{data_cell.class}"
      #         # exit if row_idx>8
      #         if data_cell.is_a? DateTime
      #           c = s.add_cell(row_idx+1, col_idx)
      #           c.set_number_format('yyyy-mm-dd')
      #           c.change_contents(data_cell)
      #         else
      #           s.add_cell(row_idx+1, col_idx, data_cell)
      #         end
      #
      #       end
      #     end
      #   end
      #   merged_book.write filename
      # end
      puts "finished!"
    end

    private
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
