require 'spreadsheet' # https://github.com/zdavatz/spreadsheet/blob/master/GUIDE.md
require 'rubyXL'      # https://github.com/weshatheleopard/rubyXL
require 'fileutils'
require 'date'
require 'byebug'

require "merge_excel/wbook"
require "merge_excel/sheets"
require "merge_excel/sheet"
require "merge_excel/header"
require "merge_excel/data_row"
require "merge_excel/version"
require "merge_excel/excel"
require "merge_excel/format_detector"
require "merge_excel/spreadsheet_parser"
require "merge_excel/spreadsheet_writer"
require "merge_excel/extensions/spreadsheet/worksheet"
require "merge_excel/extensions/spreadsheet/workbook"
require "merge_excel/extensions/rubyXL/worksheet"
require "merge_excel/extensions/rubyXL/workbook"

require "merge_excel/settings/parser"
require "merge_excel/settings/data_row"
require "merge_excel/settings/data_rows"
require "merge_excel/settings/extra_data_row"
require "merge_excel/settings/extra_data_rows"
