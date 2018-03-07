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




## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/AEEGSI/merge_excel. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](contributor-covenant.org) code of conduct.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).
