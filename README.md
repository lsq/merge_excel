# MergeExcel

This Ruby library has been created in order to merge multiple Excel files (both '.xls' and '.xlsx') into one.
To experiment with that code, run `bin/console` for an interactive prompt.


## Installation

Add this line to your application's Gemfile:

```ruby
gem 'merge_excel'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install merge_excel

## Usage

Given `excel_files_dir` the directory where the Excel files you want to merge are placed and `merged_file_path` the path or filename of the destination Excel file.


```ruby
require 'merge_excel'

me = MergeExcel::Excel.new(excel_files_dir)
me.merge merged_file_path
```
By default the command will merge:
- all worksheets
- the header line (first row)
- all rows until an empty one
- all columns

### Options

Options allows you to control what to merge.

```ruby
options = {
  selector_pattern: "*.xls",
  sheets_idxs:      [1],
  data_rows: {
    1 => {
      header_row: 0, # or :missing
      data_starting_row: 1,
      data_first_column: 0,
      data_last_column: 5 # omit or :last to get all columns
    }
  }
}

me = MergeExcel::Excel.new(excel_files_dir, options)
me.merge merged_file_path
```

In this example will be merged only '.xls' files, of which only the columns from 0 to 5 of the second sheet.

#### Options explanation

Note that all indexes (sheets, rows and columns) are 0-based.

- `selector_pattern` is used by `Dir.glob` method for filtering files inside the source directory. Default: `"*.{xls|xlsx}"`.

- `sheets_idxs` permit to select the sheets to import/merge. Pass an array with indexes (0-based) or `:all`. Default: `:all`.

- pass an `{ sheet_index => rules }` hash to `data_rows` in order to specify which rows and column you want to import. You can pass `:any` in place of the `sheet_index` for default rules.
Rules:
    - `header_row` set the header row index. Put `:missing` to omit the header. Default is `0`
    - `data_starting_row` set the index of the **first row** to import, default is `1`
    - `data_first_column` set the index of the **first column** to import, default is `0`
    - `data_last_column` set the index of the **last column** to import, default is `:last`

- pass an `{ sheet_index => rules }` hash to `extra_columns` in order to specify extra columns you want to be added to the final merge (for example you want to insert the filename you get the data). Rules:
  - `position` define the column position of extra data, you can put at `:beginning` or  `:end`
  - `data` define an array of desired extra data
    - `type` can be `:filename` if you want to add filename info or `:cell_value` if you want to pick extra info from a cell
    - `label` allow you to set extra data header name
    - when you select `:cell_value` you have to specify `sheet_idx`, `row_idx` and `col_idx`.

Another example:

```ruby
options = {
  sheets_idxs:  :all,
  data_rows: {
    1 => {
      header_row: 0,
      data_starting_row: 1,
      data_first_column: 0 # all columns
    },
    :any => {
      header_row: :missing,
      data_starting_row: 1,
      data_first_column: 0,
      data_last_column: 5
    }
  },
  extra_columns: {
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
        }
      ]
    }
  }
}
```


## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/AEEGSI/merge_excel. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](contributor-covenant.org) code of conduct.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).
