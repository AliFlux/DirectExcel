# DirectExcel
#### Easily read and write into MS Excel spreadsheets

![DirectExcel](https://raw.github.com/AliFlux/DirectExcel/master/examples/images/directexcel.png)

## Introduction

"DirectExcel" is an easy way to read and write spreadsheets in pure PHP.

There are dozens of libraries for reading/writing spreadsheets but they are far too complicated. DirectExcel makes things simpler by providing just 2 functions for reading and writing data.

## Features

- Tiny size
- Single class file
- Zero dependencies
- Object-oriented usage
- Easy to use (just 2 functions)
- One function to create and download spreadsheets
- Simply works with *arrays*

## Installation

Just copy the [class.directexcel.php](class.directexcel.php) and the [fonts directory](fonts) into your working path, where it is easily accessible by your PHP script.

## Read data

```php
<?php

require_once("class.directexcel.php");

// single function read
$data = DirectExcel::readArray("inventions.xlsx");

header("content-type: text/plain");
echo "Reading inventions.xlsx\n\n";

// show the data
print_r($data);

?>
```

## Write data

```php
<?php

require_once("class.directexcel.php");

// sample data is in array
$data = array(
	"Aircrafts" => array(
		array(
			"Name" => "SR-71 BlackBird",
			"Speed" => "Mach 3+",
		),
		array(
			"Name" => "F-22 Raptor",
			"Speed" => "Mach 1.82",
		),
	),
	"Aircraft Carrier" => array(
		array(
			"Name" => "USS Nimitz",
			"Classification" => "CVN-68",
		),
		array(
			"Name" => "USS Gerald R. Ford",
			"Classification" => "CVN-78",
		),
		array(
			"Name" => "USS Ronald Reagan",
			"Classification" => "CVN-76",
		),
	)
);

// single function write
DirectExcel::write($data, "firepower.xlsx");

header("content-type: text/plain");
echo "Excel file written successfully: \nfirepower.xlsx";

?>
```

You'll find plenty more to play with in the [examples](examples/) folder.

## Data format

To make things simpler, this class was created to work with arrays. The array format for reading and writing is...

```php
<?php
$data = array(
	"Worksheet1" => array(
		array(
			"Column1" => "Cell 1 information",
			"Column2" => "Cell 2 information",
			// more columns...
		),
		// more rows...
	),
	// more worksheets...
);
?>
```

## Problems with CSV

CSVs are usually used as an easy way to read/write into excel format. But CSV has a lot of inherent issues.

- When CSVs get exported into a spreadsheet software, their columns are shrinked.
- Multiline cells look unreadable.
- Phone numbers in cells get messed up. (e.g. +92-353-5338954 becomes -5338968).
- Only one worksheet is supported per CSV file.
- And it's not professional enough if you're using it on your site.

## Note

- Currently this class only supports the modern *.xlsx* format. The legacy *.xls* format is not supported.
- This class was created during a 6 hour hackathon. A lot of testing is yet to be done.

## Contributing

Please submit bug reports, suggestions and pull requests to the [GitHub issue tracker](https://github.com/AliFlux/DirectExcel/issues).
I would be happy if someone helps in its testing and extending its functionalities.

## License

This software is licenced under the [LGPL 2.1](http://www.gnu.org/licenses/lgpl-2.1.html). Please read LICENSE for information on the
software availability and distribution.
