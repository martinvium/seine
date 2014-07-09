Seine
=====

[![Build Status](https://travis-ci.org/martinvium/seine.svg)](https://travis-ci.org/martinvium/seine)

Write spreadsheets of various formats to a stream

Low memory, high performance library for writing large spreadsheets in various standard formats. 
Only a small subset of features are included, which means row level styling and right now, no 
formulas either.

Memory: Everything is written to disk (stream), so memory overhead is pretty much zero.
Speed:  It's pretty damn fast! However because of the memory constraint, it's not mindblowing.

Stability
---------

This library is ALPHA/BETA quality.

Dependencies
------------

* PHP 5.3
* ZipArchive (only OOXML and you can create your own zip compressor, if you prefer another solution)

Writers
-------

* Office Open XML Spreadsheet (.xlsx)
* Microsoft Excel 2003 XML (.xml)
* CSV (.csv)

Examples
--------

Create a new document and close it after you done.

```php
<?php
use Seine\Seine;

$seine = new Seine(array('writer' => 'ooxml2007')); // writer options are: csv, ooxml2007, oxml2003
$doc = $seine->newDocument('example.xlsx');

// Add rows and styles

$doc->close();
?>
```

Create a new document using an existing stream

```php
<?php
$fp = fopen('filename.csv', 'w');
$doc = $seine->newDocumentFromStream($fp);
$doc->close();
fclose($fp);
?>
```

Add 100.000 rows with 25 columns in ~25 seconds.

```php
<?php
$sheet = $doc->newSheet();
for($i = 0; $i < 100000; $i++) {
    $sheet->addRow(range(0, 25));
}
?>
```

Add styling to a row.

```php
<?php
$style = $doc->newStyle()
             ->setFontBold(true)
             ->setFontFamily('Aria')
             ->setFontSize('14');
$row = $seine->getRow(array('cell1', 'cell2'));
$row->addStyle($style);
?>
```
