module MergeExcel
  class WBook
    attr_reader :sheets, :settings

    def initialize(settings)
      @sheets   = Sheets.new
      @settings = settings
    end

    def import_data(filepath)
      book     = SpreadsheetParser.open(filepath)
      filename = File.basename(filepath)

      book.worksheets.each_with_index do |sheet, idx|
        if @settings.read_sheet? idx
          sh_stt  = @settings.data_rows.get(idx)
          sh      = select_sheet(idx) || create_sheet(sheet, idx, sh_stt)
          row_idx = sh_stt.data_starting_row
          loop do
            break unless cells = sheet.row_values(row_idx, sh_stt.data_first_column, sh_stt.data_last_column)
            sh.add_data_row(cells, @settings.extra_columns.get(idx), book)
            row_idx+=1
          end
        end
      end
    end

    def export(filepath)
      print "Exporting data..."
      SpreadsheetWriter.create(filepath) do |wb|
        @sheets.each_with_index do |sheet, sheet_idx|
          s = wb.add_worksheet_at(sheet_idx, sheet.name)
          if sheet.header
            s.add_row_at(0, sheet.header)
            offset = 1
          else
            offset = 0
          end
          sheet.data_rows.each_with_index do |data_row, row_idx|
            s.add_row_at(row_idx+offset, data_row)
          end
        end
      end
      puts "finished!"
    end


    private

    def select_sheet(idx)
      @sheets.get_sheet(original_idx: idx)
    end

    def create_sheet(sheet, idx, sh_stt)
      sh = Sheet.new(sheet, idx)
      if sh_stt.import_header?
        arr = sheet.row_values(sh_stt.header_row, sh_stt.data_first_column, sh_stt.data_last_column)
        sh.add_header(arr, @settings.extra_columns.get(idx))
      end
      @sheets << sh
      sh
    end

  end
end
