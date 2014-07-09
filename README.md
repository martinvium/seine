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

This is still only ALPHA/BETA quality.

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

Create a configured Book and close after you done.

    ``` php
    use Seine\Parser\DOM\DOMFactory;
    use Seine\Configuration;

    $fp = fopen('example.xlsx', 'w');
    $config = new Configuration();
    $config->setWriter(Configuration::WRITER, 'OfficeOpenXml2007StreamWriter');
    $factory = new DOMFactory($config);
    $book = $factory->getConfiguredBook();

    // Add rows and styles

    $book->close();
    fclose($fp);
    ```

Add 100.000 rows with 25 columns in ~25 seconds.

    ``` php
    $sheet = $book->newSheet();
    for($i = 0; $i < 100000; $i++) {
        $sheet->addRow($this->factory->getRow(range(0, 25)));
    }
    ```

Add styling to a Row.

    ``` php
    $style = $book->newStyle()
                  ->setFontBold(true)
                  ->setFontFamily('Aria')
                  ->setFontSize('14');
    $row = $factory->getRow(array('cell1', 'cell2'));
    $row->addStyle($style);
    ```

Custom Compressor for OOXML

    ``` php
    class CommandLineCompressor implements Seine\Compressor
    {
        public function compressDir($source, $destination)
        {
            exec('zip ' . escapeshellarg($source) . escapeshellarg($destination)); // Most likely does not work :)
        }
    }

    $compressor = new CommandLineCompressor;
    $writer = new Seine\Writer\OfficeOpenXML2007StreamWriter($fp, $compressor);
    ```