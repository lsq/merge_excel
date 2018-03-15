require 'spec_helper'
require 'helpers'

RSpec.configure do |c|
  c.include Helpers
end


describe MergeExcel do
  it 'has a version number' do
    expect(MergeExcel::VERSION).not_to be nil
  end

  it 'select only xls files' do
    options = {
      selector_pattern: "*.xls"
    }
    e = MergeExcel::Excel.new(mixed_xls_xlsx_dir_path, options)
    expect(e.files.size).to eq 2
  end

  it 'select only xlsx files' do
    options = {
      selector_pattern: "*.xlsx"
    }
    e = MergeExcel::Excel.new(mixed_xls_xlsx_dir_path, options)
    expect(e.files.size).to eq 3
  end

  it 'select both xls and xlsx files' do
    options = {
      selector_pattern: "*.{xlsx,xls}"
    }
    e = MergeExcel::Excel.new(mixed_xls_xlsx_dir_path, options)
    expect(e.files.size).to eq 5

    e2 = MergeExcel::Excel.new(mixed_xls_xlsx_dir_path, {}) # default version
    expect(e2.files.size).to eq 5
  end




  context 'when merge xlsx files' do
    let(:suffix) { "xlsx" }

    it "skip the header line" do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      options = {
        sheets_idxs:      [1],
        data_rows: {
          1 => {
            header_row: :missing, # or :missing
            data_starting_row: 1
          }
        }
      }
      e = MergeExcel::Excel.new(xlsx_dir2_path, options)
      e.merge output_file_path
      m_book = RubyXL::Parser.parse(output_file_path)
      expect(m_book.cell_value_at(0,0,2)).to eq "Ruby Patel"
    end


    it "skip two data lines" do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      options = {
        sheets_idxs: [1],
        data_rows: {
          1 => {
            header_row: 0, # or :missing
            data_starting_row: 3
          }
        }
      }
      e = MergeExcel::Excel.new(xlsx_dir2_path, options)
      e.merge output_file_path
      m_book = RubyXL::Parser.parse(output_file_path)
      expect(m_book.cell_value_at(0,1,2)).to eq "Devin Huddleston"
      expect(m_book.cell_value_at(0,5,2)).to eq "Winnie Moss"
    end


    it "take column 2 and 3" do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      options = {
        sheets_idxs:      [1],
        data_rows: {
          1 => {
            header_row: 0, # or :missing
            data_starting_row: 1,
            data_first_column: 2,
            data_last_column: 3 # omit or :last to get all columns
          }
        }
      }
      e = MergeExcel::Excel.new(xlsx_dir2_path, options)
      e.merge output_file_path
      m_book = RubyXL::Parser.parse(output_file_path)
      expect(m_book.cell_value_at(0,1,0)).to eq "Ruby Patel"
      expect(m_book.cell_value_at(0,1,1)).to eq "Stockholm"
      expect(m_book.cell_value_at(0,11,0)).to eq "Winnie Moss"
      expect(m_book.cell_value_at(0,11,1)).to eq "Barcelona"
    end


    it "put filename at end" do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      options = {
        sheets_idxs:      [1],
        data_rows: {
          1 => {
            header_row: 0,
            data_starting_row: 1
          }
        },
        extra_columns: {
          any: {
            position: :end,
            data: [
              {
                type: :filename,
                label: "Filename"
              }
            ]
          }
        }
      }
      e = MergeExcel::Excel.new(xlsx_dir2_path, options)
      e.merge output_file_path
      m_book = RubyXL::Parser.parse(output_file_path)
      expect(m_book.cell_value_at(0,0,0)).to eq "Order ID"
      expect(m_book.cell_value_at(0,0,5)).to eq "Filename"
      expect(m_book.cell_value_at(0,1,5)).to eq "wbook1.#{suffix}"
      expect(m_book.cell_value_at(0,11,5)).to eq "wbook3.#{suffix}"
    end


    [
      { str: "one", num: 1, arr: [1] },
      { str: "two", num: 2, arr: [0,2] },
      { str: "three", num: 3, arr: [0,1,2] },
      { str: "all", num: 3, arr: :all }
    ].each do |h|
      it "merge #{h[:str]} sheet(s)" do
        output_file_path = File.join(results_dir_path, "merged.#{suffix}")
        e = MergeExcel::Excel.new(xlsx_dir_path, { sheets_idxs: h[:arr] })
        e.merge output_file_path
        m_book = RubyXL::Parser.parse(output_file_path)
        expect(m_book.count_worksheets).to eq h[:num]
      end
    end




    it 'Two sheets specifying sheet_idx, header_rows, data_rows and extra_columns' do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      e = MergeExcel::Excel.new(xlsx_dir2_path, options_2)
      e.merge output_file_path

      m_book = RubyXL::Parser.parse(output_file_path)
      expect(m_book.cell_value_at(0,0,0)).to eq "Filename"
      expect(m_book.cell_value_at(0,0,1)).to eq "Company"
      expect(m_book.cell_value_at(0,0,2)).to eq "Amount"
      expect(m_book.cell_value_at(0,0,3)).to eq "Order ID"
      expect(m_book.cell_value_at(0,0,7)).to eq "Country"
      expect(m_book.cell_value_at(0,0,8)).to eq nil

      expect(m_book.cell_value_at(0,1,0)).to eq "wbook1.#{suffix}"
      expect(m_book.cell_value_at(0,1,1)).to eq "ACME"
      expect(m_book.cell_value_at(0,1,2)).to eq 45
      expect(m_book.cell_value_at(0,1,7)).to eq "Sweden"
      expect(m_book.cell_value_at(0,1,8)).to eq nil

      expect(m_book.cell_value_at(0,4,0)).to eq "wbook2.#{suffix}"
      expect(m_book.cell_value_at(0,4,1)).to eq "Butugah"
      expect(m_book.cell_value_at(0,4,2)).to eq nil
      expect(m_book.cell_value_at(0,4,7)).to eq "France"
      expect(m_book.cell_value_at(0,4,8)).to eq nil

      expect(m_book.cell_value_at(0,11,0)).to eq "wbook3.#{suffix}"
      expect(m_book.cell_value_at(0,11,1)).to eq "ZZZKL"
      expect(m_book.cell_value_at(0,11,2)).to eq 29
      expect(m_book.cell_value_at(0,11,7)).to eq "Spain"
      expect(m_book.cell_value_at(0,11,8)).to eq nil

      expect(m_book.cell_value_at(1,0,0)).to eq "Filename"
      expect(m_book.cell_value_at(1,0,1)).to eq "Company"
      expect(m_book.cell_value_at(1,0,2)).to eq "Amount"
      expect(m_book.cell_value_at(1,0,3)).to eq "Order ID"
      expect(m_book.cell_value_at(1,0,6)).to eq "Sales"
      expect(m_book.cell_value_at(1,0,7)).to eq nil

      expect(m_book[0][12]).to eq nil
    end



    it 'All sheets specifying sheet_idx, header_rows and data_rows' do
      output_file_path = File.join(results_dir_path, "merged.xlsx")
      e = MergeExcel::Excel.new(xlsx_dir_path, options_1)
      e.merge output_file_path

      m_book = RubyXL::Parser.parse(output_file_path)
      row_indexes = {0 => 0, 1 => 0, 2 => 0}
      [1,2,3].each do |original_wb_num|
        name = "wbook#{original_wb_num}.xlsx"
        o_book = RubyXL::Parser.parse(File.join(xlsx_dir_path, name))

        [0,1,2].each do |sheet_idx|
          m_sheet = m_book[sheet_idx]
          o_sheet = o_book[sheet_idx]
          i = original_wb_num==1 ? 0 : 1
          while o_row=o_sheet[i] # not nil
            o_values = o_row.cells.map{|c| c && c.value}
            if o_values.empty? || o_values.all?(&:nil?)
              i+=1
              next
            end
            m_row = m_sheet[row_indexes[sheet_idx]]
            m_values = m_row && m_row.cells.map{|c| c && c.value}
            o_values.each_with_index do |v, idx|
              expect(m_values[idx]).to eq o_values[idx]
            end
            i += 1
            row_indexes[sheet_idx] += 1
          end
        end
      end
    end
  end





  context 'when merge xls files' do
    let(:suffix) { "xls" }


    it "skip the header line" do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      options = {
        sheets_idxs:      [1],
        data_rows: {
          1 => {
            header_row: :missing,
            data_starting_row: 1
          }
        }
      }
      e = MergeExcel::Excel.new(xls_dir2_path, options)
      e.merge output_file_path
      m_book = Spreadsheet.open(output_file_path)
      expect(m_book.cell_value_at(0,0,2)).to eq "Ruby Patel"
    end


    it "skip two data lines" do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      options = {
        sheets_idxs: [1],
        data_rows: {
          1 => {
            header_row: 0, # or :missing
            data_starting_row: 3
          }
        }
      }
      e = MergeExcel::Excel.new(xls_dir2_path, options)
      e.merge output_file_path
      m_book = Spreadsheet.open(output_file_path)
      expect(m_book.cell_value_at(0,1,2)).to eq "Devin Huddleston"
      expect(m_book.cell_value_at(0,5,2)).to eq "Winnie Moss"
    end


    it "take column 2 and 3" do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      options = {
        sheets_idxs: [1],
        data_rows: {
          1 => {
            header_row: 0,
            data_starting_row: 1,
            data_first_column: 2,
            data_last_column: 3 # omit or :last to get all columns
          }
        }
      }
      e = MergeExcel::Excel.new(xls_dir2_path, options)
      e.merge output_file_path
      m_book = Spreadsheet.open(output_file_path)
      expect(m_book.cell_value_at(0,1,0)).to eq "Ruby Patel"
      expect(m_book.cell_value_at(0,1,1)).to eq "Stockholm"
      expect(m_book.cell_value_at(0,11,0)).to eq "Winnie Moss"
      expect(m_book.cell_value_at(0,11,1)).to eq "Barcelona"
    end


    it "put filename at end" do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      options = {
        sheets_idxs: [1],
        data_rows: {
          1 => {
            header_row: 0,
            data_starting_row: 1
          }
        },
        extra_columns: {
          any: {
            position: :end,
            data: [
              {
                type: :filename,
                label: "Filename"
              }
            ]
          }
        }
      }
      e = MergeExcel::Excel.new(xls_dir2_path, options)
      e.merge output_file_path
      m_book = Spreadsheet.open(output_file_path)
      expect(m_book.cell_value_at(0,0,0)).to eq "Order ID"
      expect(m_book.cell_value_at(0,0,5)).to eq "Filename"
      expect(m_book.cell_value_at(0,1,5)).to eq "wbook1.#{suffix}"
      expect(m_book.cell_value_at(0,11,5)).to eq "wbook3.#{suffix}"
    end


    [
      { str: "one", num: 1, arr: [1] },
      { str: "two", num: 2, arr: [0,2] },
      { str: "three", num: 3, arr: [0,1,2] },
      { str: "all", num: 3, arr: :all }
    ].each do |h|
      it "merge #{h[:str]} sheet(s)" do
        output_file_path = File.join(results_dir_path, "merged.#{suffix}")
        e = MergeExcel::Excel.new(xlsx_dir_path, { sheets_idxs: h[:arr] })
        e.merge output_file_path
        m_book = Spreadsheet.open(output_file_path)
        expect(m_book.count_worksheets).to eq h[:num]
      end
    end


    it 'Two sheets specifying sheet_idx, header_rows, data_rows and extra_columns' do
      output_file_path = File.join(results_dir_path, "merged.#{suffix}")
      e = MergeExcel::Excel.new(xls_dir2_path, options_2)
      e.merge output_file_path

      m_book = Spreadsheet.open(output_file_path)

      expect(m_book.cell_value_at(0,0,0)).to eq "Filename"
      expect(m_book.cell_value_at(0,0,1)).to eq "Company"
      expect(m_book.cell_value_at(0,0,2)).to eq "Amount"
      expect(m_book.cell_value_at(0,0,3)).to eq "Order ID"
      expect(m_book.cell_value_at(0,0,7)).to eq "Country"
      expect(m_book.cell_value_at(0,0,8)).to eq nil

      expect(m_book.cell_value_at(0,1,0)).to eq "wbook1.#{suffix}"
      expect(m_book.cell_value_at(0,1,1)).to eq "ACME"
      expect(m_book.cell_value_at(0,1,2)).to eq 45
      expect(m_book.cell_value_at(0,1,7)).to eq "Sweden"
      expect(m_book.cell_value_at(0,1,8)).to eq nil

      expect(m_book.cell_value_at(0,4,0)).to eq "wbook2.#{suffix}"
      expect(m_book.cell_value_at(0,4,1)).to eq "Butugah"
      expect(m_book.cell_value_at(0,4,2)).to eq nil
      expect(m_book.cell_value_at(0,4,7)).to eq "France"
      expect(m_book.cell_value_at(0,4,8)).to eq nil

      expect(m_book.cell_value_at(0,11,0)).to eq "wbook3.#{suffix}"
      expect(m_book.cell_value_at(0,11,1)).to eq "ZZZKL"
      expect(m_book.cell_value_at(0,11,2)).to eq 29
      expect(m_book.cell_value_at(0,11,7)).to eq "Spain"
      expect(m_book.cell_value_at(0,11,8)).to eq nil

      expect(m_book.cell_value_at(1,0,0)).to eq "Filename"
      expect(m_book.cell_value_at(1,0,1)).to eq "Company"
      expect(m_book.cell_value_at(1,0,2)).to eq "Amount"
      expect(m_book.cell_value_at(1,0,3)).to eq "Order ID"
      expect(m_book.cell_value_at(1,0,6)).to eq "Sales"
      expect(m_book.cell_value_at(1,0,7)).to eq nil

      expect(m_book.worksheet(0).row(12)).to eq []
    end


    it 'merge all sheets when specify sheet_idx, header_rows and data_rows' do
      output_file_path = File.join(results_dir_path, "merged.xls")
      e = MergeExcel::Excel.new(xls_dir_path, options_1)
      e.merge output_file_path

      m_book = Spreadsheet.open(output_file_path)
      row_indexes = {0 => 0, 1 => 0, 2 => 0}
      [1,2,3].each do |original_wb_num|
        name = "wbook#{original_wb_num}.xls"
        o_book = Spreadsheet.open(File.join(xls_dir_path, name))

        [0,1,2].each do |sheet_idx|
          m_sheet = m_book.worksheet(sheet_idx)
          o_sheet = o_book.worksheet(sheet_idx)
          i = original_wb_num==1 ? 0 : 1
          while (o_row=o_sheet.row(i)).empty? # not nil
            o_values = o_row#.map{|c| c && c.value}
            if o_values.empty? || o_values.all?(&:nil?)
              i+=1
              next
            end
            m_row = m_sheet.row(row_indexes[sheet_idx])
            m_values = m_row #&& m_row.map{|c| c && c.value}
            o_values.each_with_index do |v, idx|
              expect(m_values[idx]).to eq o_values[idx]
            end
            i += 1
            row_indexes[sheet_idx] += 1
          end
        end
      end
    end
  end
end
