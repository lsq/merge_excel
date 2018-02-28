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
      header_rows: {
        0 => {
          row_idx: 0,
          from_col_idx: 0,
          to_col_idx: 9
        },
        1 => {
          row_idx: 0,
          from_col_idx: 0,
          to_col_idx: 7
        },
        2 => {
          row_idx: 0,
          from_col_idx: 0,
          to_col_idx: 2
        }
      },
      data_rows: {
        0 => {
          first_row_idx: 1,
          from_col_idx: 0,
          to_col_idx: 9
        },
        1 => {
          first_row_idx: 1,
          from_col_idx: 0,
          to_col_idx: 7
        },
        2 => {
          first_row_idx: 1,
          from_col_idx: 0,
          to_col_idx: 2
        }
      }
      # extra_data: {
      #   0 => {
      #     position: :beginning,
      #     data: [
      #       {
      #         heading_text: "Filename",
      #         data: :filename
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
