# xlsx-readable
XLSX Readable Stream for CommonJS

## Motivation
A number of XLSX libraries already exist for the CommonJS standard. This library exists for a specific purpose: To enable high-speed 'scraping' of data from XLSX files. Other projects typically attempt to support cell formatting, random-access, and/or present data as key-value pairs. This is not conducive to projects where the only need is to collect data for input into import or transfer algorithms, especially in cases where the workbook may have been formatted for human consumption (rather than in a machine-friendly layout.) This library exposes the workbook as a real-time readable stream, and makes no assumptions about the location or format of column headers (nor that they exist at all.)

_Note that no native Excel format is readily suited for streaming (as they have cross-dependencies between data sets.) This library attempts to provide a streamable interface as much as can be provided. In particular, there may be noticable delay before worksheets become available, due to overhead from ZIP parsing and reading in the shared string table._

## Installation
`npm install xlsx-readable`

## API Reference
Import the library:
`let XlsxReadable = require("../../xlsx-readable");`

####Build your open options:
`options = {
 ignoreEmpty: boolean, // Skip empty rows, or pass through as an empty array?
 sheets: [array_of_strings] // Names of sheets to stream through.
}`

####Open a file as your stream source:
`let xlsx = new XlsxReadable(file_path, options);`

####Events emitted by XlsxReadable:
`"worksheet" // A worksheet you requested is ready to stream. It may be used immediately.
"finish" // Done finding worksheets.
"error" // An error was reported during parsing of file.`

####Worksheet Event:
Passes an argument object that exposes an integer 'index', reporting to the base-1 position of the worksheet, and a 'openReadStream' function, which opens a standard 'Readable' stream for that worksheet's rows.
`let readable = worksheet.openReadStream()`

####Worksheet Readable
Behavior is identical to a Readable stream in 'object' mode. Each data object is an array containing one row of sheet data.

####Utility Method: objectify()
Data objects are bare arrays. This method is provided to allow flexibility in finding (or defining) your own header row. Converts bare row into object based on key-value pairs.
`let row = xlsx.objectify(header_key_row, data_value_row)`

## Contributors
Improvements and expansions of features welcome (keeping in mind the overall objective of this library.) I don't have a lot of available time to manage PR's, but will definitely consider significant ones when I can.
