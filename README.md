# Minimalist ODS

`minimalist_ods` is a minimalist gem for generating ODS (Open Document Spreadsheet) files in Ruby. This gem is designed specifically for exporting data to ODS format.

## Installation

Add this line to your Gemfile:

```ruby
gem 'minimalist_ods'
```

And then execute:

```sh
bundle install
```

Or install it yourself as:

```sh
gem install minimalist_ods
```

## Usage

### Create an ODS File

To create an ODS file, first initialize an `ODS` object with the file name and the creator's name:

```ruby
require 'minimalist_ods'

# by default, creator metadata will be minimalist-ods, you can set it as the second argument in your constructor
ods = MinimalistODS.new('example.ods', 'Creator')
```

### Open a Table

To add a table, use the `open_table` method with the table name and the number of columns:

```ruby
ods.open_table('Sheet1', 3)
```

### Add Rows

To add a row, use the `add_row` method with an array of values:

```ruby
ods.add_row(['Name', 'Age', 'City'])
ods.add_row(['Alice', 30, 'New York'])
ods.add_row(['Bob', 25, 'San Francisco'])
```

### Close the Table

Once you have finished adding rows, close the table using:

```ruby
ods.close_table
```

### Close the File

Finally, close the file to save the changes:

```ruby
ods.close_file
```

### Complete Example

```ruby
require 'minimalist_ods'

ods = MinimalistODS.new('example.ods', 'Your Name')
ods.open_table('Sheet1', 3)
ods.add_row(['Name', 'Age', 'City'])
ods.add_row(['Alice', 30, 'New York'])
ods.add_row(['Bob', 25, 'San Francisco'])
ods.close_table
ods.close_file
```

### Create ODS in buffer

It is possible to avoid writing the file to disk if you pass a StringIO object as the first parameter. You will still need to close the file before reading the buffer.

```ruby
file_buffer = StringIO.new
ods = MinimalistODS.new(file_buffer)
...
ods.close_file
# Do something with file_buffer.read
```

## Contributions

Contributions are welcome. Please open an issue or a pull request on the [GitHub repository](https://github.com/tachomex/minimalist_ods).

## License

This gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).
