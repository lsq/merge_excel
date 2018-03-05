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
      input_dir: mixed_xls_xlsx_dir_path,
      output_file_path: File.join(results_dir_path, "fake_name.xlsx")
    }
    e = MergeExcel::Excel.new(options, "*.xls")
    expect(e.files.size).to eq 2
  end

  it 'select only xlsx files' do
    options = {
      input_dir: mixed_xls_xlsx_dir_path,
      output_file_path: File.join(results_dir_path, "fake_name.xlsx")
    }
    e = MergeExcel::Excel.new(options, "*.xlsx")
    expect(e.files.size).to eq 3
  end

  it 'select both xls and xlsx files' do
    options = {
      input_dir: mixed_xls_xlsx_dir_path,
      output_file_path: File.join(results_dir_path, "fake_name.xlsx")
    }
    e = MergeExcel::Excel.new(options, "*.{xlsx,xls}")
    expect(e.files.size).to eq 5

    e2 = MergeExcel::Excel.new(options) # default version
    expect(e2.files.size).to eq 5
  end


  it 'merge all sheets when specify sheet_idx, header_rows and data_rows (xlsx)' do
    output_file_path = File.join(results_dir_path, "merged.xlsx")
    options = options_1(xlsx_dir_path, output_file_path)
    e = MergeExcel::Excel.new(options, "*.{xlsx,xls}")
    e.merge

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


  it 'merge all sheets when specify sheet_idx, header_rows and data_rows (xls)' do
    output_file_path = File.join(results_dir_path, "merged.xls")
    options = options_1(xls_dir_path, output_file_path)
    e = MergeExcel::Excel.new(options, "*.{xlsx,xls}")
    e.merge

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
