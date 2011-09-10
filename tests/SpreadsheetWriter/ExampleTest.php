<?php
namespace SpreadSheetWriter;

require_once(dirname(__DIR__) . '/bootstrap.php');

use SpreadSheetWriter\Parser\DOM\DOMFactory;

class ExampleTest extends \PHPUnit_Framework_TestCase
{
    public function testMultipleSheetsWithStyles()
    {
        $actual_file = __DIR__ . '/_files/out.xml';
        
        $factory = new DOMFactory();
        $book = $factory->getBook();
        $fp = fopen($actual_file, 'w');
        $book->setWriter($factory->getWriterFactory()->getOfficeXML2003StreamWriter($fp));

        $sheet = $book->addSheetByName('more1');
        for($i = 0; $i < 100; $i++) {
            $sheet->addRow($factory->getRow(array(
                'cell1',
                'cell2',
                'cell3',
                'cell4'
            )));
        }

        $sheet = $book->addSheetByName('more2');
        $style = $sheet->addStyleById('mystyle')->setFontBold(true);
        $sheet->addRow($factory->getRow(array('head1', 'head2', 'head3', 'head4'))->setStyle($style));
        for($i = 0; $i < 1000; $i++) {
            $sheet->addRow($factory->getRow(array(
                'cell1',
                'cell2',
                rand(100, 10000),
                'cell4'
            )));
        }
        $book->close();
        fclose($fp);
        
        $actual = file_get_contents($actual_file);
        $this->assertSelectCount('Workbook Styles Style', 2, $actual, 'default style + one custom style');
        $this->assertSelectCount('Workbook Worksheet', 2, $actual, 'total number of sheets');
        $this->assertSelectCount('Workbook Worksheet Table Row', 1101, $actual, 'total number of rows');
    }
    
    public function testLargeNumberOfRowsDoesNotImpedeMemory()
    {
        $memory_limit = 20 * 1000 * 1000; // 20 MB
        $time_limit_seconds = 3.0; // 3 seconds
        $num_rows = 10000;
        $num_columns = 100;
        
        $start_timestamp = microtime(true);
        $actual_file = __DIR__ . '/_files/out_perf.xml';
        $factory = new DOMFactory();
        $book = $factory->getBook();
        
        $fp = fopen($actual_file, 'w');
        $book->setWriter($factory->getWriterFactory()->getOfficeXML2003StreamWriter($fp));
        $sheet = $book->addSheetByName('more1');
        for($i = 0; $i < $num_rows; $i++) {
            $sheet->addRow($factory->getRow(range(0, $num_columns)));
        }
        $book->close();
        fclose($fp);
        
        $this->assertLessThan($memory_limit, memory_get_peak_usage(true), 'memory limit reached');
        $this->assertLessThan($time_limit_seconds, (microtime(true) - $start_timestamp), 'time limit reached');
    }
}