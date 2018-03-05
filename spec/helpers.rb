module Helpers

  def mixed_xls_xlsx_dir_path
    File.join(__dir__, "mixed_xls_xlsx")
  end

  def xlsx_dir_path
    File.join(__dir__, "xlsx_files_1")
  end

  def xls_dir_path
    File.join(__dir__, "xls_files_1")
  end

  def results_dir_path
    File.join(__dir__, "results")
  end



  def options_1(input_dir_path, output_file_path)
    options = {
      input_dir:        input_dir_path,
      output_file_path: output_file_path,
      sheets_idxs:      [0,1,2],
      data_rows: {
        0 => {
          header_row: 0, # or :missing
          data_starting_row: 1,
          data_first_column: 0,
          data_last_column: 9 # omit or :last to get all columns
        },
        1 => {
          header_row: 0,
          data_starting_row: 1,
          data_first_column: 0,
          data_last_column: 7
        },
        2 => {
          header_row: 0,
          data_starting_row: 1,
          data_first_column: 0,
          data_last_column: 2
        }
      },
      extra_data_rows: {
        any: {          
        }
      }
      # extra_data: {
      #   0 => {
      #     position: :beginning,
      #     data: [
      #       {
      #         type: :filename
      #         label: "Filename",
      #       },
      #       {
      #         type: :cell_value,
      #         label: "Company",
      #         sheet_idx: 6,
      #         row_idx: 5,
      #         col_idx: 4,

      #       }
      #     ]
      #   },
      #   1 => {
      #     position: :beginning,
      #     data: [
      #       {
      #         heading_text: "Filename",
      #         data: :filename
      #       }
      #     ]
      #   },
      #   2 => {
      #     position: :beginning,
      #     data: [
      #       {
      #         heading_text: "Filename",
      #         data: :filename
      #       }
      #     ]
      #   }
      # }

    }
    options
  end
end
