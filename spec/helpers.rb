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

  def xlsx_dir2_path
    File.join(__dir__, "xlsx_files_2")
  end

  def xls_dir2_path
    File.join(__dir__, "xls_files_2")
  end



  def options_1
    {
      selector_pattern: "*.{xlsx,xls}",
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
    }
  end

  def options_2
    {
      selector_pattern: "*.{xlsx,xls}",
      sheets_idxs:      [1,2],
      data_rows: {
        :any => {
          header_row: 0,
          data_starting_row: 1,
        }
      },
      extra_data_rows: {
        any: {
          position: :beginning,
          data: [
            {
              type: :filename,
              label: "Filename"
            },
            {
              type: :cell_value,
              label: "Company",
              sheet_idx: 0,
              row_idx: 10,
              col_idx: 3,
            },
            {
              type: :cell_value,
              label: "Amount",
              sheet_idx: 0,
              row_idx: 3,
              col_idx: 5,
            }
          ]
        }
      }
    }
  end

end
