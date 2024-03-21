# Smart Excel to Csv converter

This is a smart tool for converting Excel files to CSV during some ETL processes.
You have two options how to work with this tool:

1. Manually enter source and target files with some properties in command line arguments
2. Use TOML file as config.

## Manually usage

This tool not implemented special options in command line arguments - you could specify:

* Source file path (required)
* Target file path (required)
* List of sheets, these should be parsed

## TOML Config

TOML Structure description:

~~~ TOML
{
    "file": {
        "source": {
            "path" : path to workbook,
            "sheets" : [Optional] Array of sheets names to parse
        },
        "target": {
            "path": Path to target file,
            [Optional] "separator": columns separator in target file,
            [Optional] "columns": columns names for target file
        }
    },
    "sheets": {
        "global": {         # settings will be applied to all sheets, if sheet settings not specified separatly
            [Optional] "columns" : List of columns names in source file to parse,
            [Optional] "names_ranges": List of named ranges in source file to parse,
            [Optional] "tables": List of table names in source file to parse,
            [Optional] "checks": List of global checks if source file
        },
        [Optional] "<sheet_name>" : {
            "sheet_name": Name of sheet, that should parsed in different way,
            [Optional] "path": Path to target file, if sheet have to write into another file,
            [Optional] "separator": columns separator in target file,
            [Optional] "columns" : List of columns names in source file to parse,
            [Optional] "names_ranges": List of named ranges in source file to parse,
            [Optional] "tables": List of table names in source file to parse,
            [Optional] "checks": List of global checks if source file
        }
    }
}
~~~