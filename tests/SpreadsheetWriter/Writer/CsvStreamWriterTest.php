<?php
namespace SpreadSheetWriter\Writer;

require_once(dirname(dirname(__DIR__)) . '/bootstrap.php');

use SpreadSheetWriter\Parser\DOM\DOMFactory;

class CsvStreamWriterTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var DOMFactory
     */
    private $factory;
    
    public function setUp()
    {
        parent::setUp();
        $this->factory = new DOMFactory();
    }
    
    public function testWritesInValidFormat()
    {
        $actual_file = __DIR__ . '/_files/actual_valid.csv';
        
        $fp = $this->makeStream($actual_file);
        
        $book = $this->makeBookWithWriter($fp);
        $sheet = $book->addSheetByName('more1');
        for($i = 0; $i < 10; $i++) {
            $sheet->addRow($this->factory->getRow(array(
                'cell1',
                'cell"2',
                'cæøå3',
                'cell4'
            )));
        }
        $book->close();
        
        fclose($fp);

        $this->assertFileEquals(__DIR__ . '/_files/expected_valid.csv', $actual_file);
    }
    
    public function testWriteAllowsCustomizingDelimiters()
    {
        $actual_file = __DIR__ . '/_files/actual_custom_delimiters.csv';
        
        $fp = $this->makeStream($actual_file);
        
        $book = $this->factory->getBook();
        $writer = $this->factory->getWriterFactory()->getCsvStreamWriter($fp);
        $writer->setFieldDelimiter(';');
        $writer->setTextDelimiter("'");
        $writer->setRowDelimiter('NEWROW');
        $book->setWriter($writer);
        $sheet = $book->addSheetByName('more1');
        for($i = 0; $i < 10; $i++) {
            $sheet->addRow($this->factory->getRow(array(
                'cell1',
                'ce"2',
                'cæøå3',
                'ce\'4'
            )));
        }
        $book->close();
        
        fclose($fp);

        $this->assertFileEquals(__DIR__ . '/_files/expected_custom_delimiters.csv', $actual_file);
    }
    
    public function testWriteSpeedAndMemoryUsage()
    {
        $memory_limit = 20 * 1000 * 1000; // 20 MB
        $time_limit_seconds = 2.0; // 2 seconds
        $num_rows = 10000;
        $num_columns = 100;
        
        $start_timestamp = microtime(true);
        $actual_file = __DIR__ . '/_files/performance.csv';
        
        $fp = $this->makeStream($actual_file);
        
        $book = $this->makeBookWithWriter($fp);
        $sheet = $book->addSheetByName('more1');
        for($i = 0; $i < $num_rows; $i++) {
            $sheet->addRow($this->factory->getRow(range(0, $num_columns)));
        }
        $book->close();
        fclose($fp);
        
        $this->assertLessThan($memory_limit, memory_get_peak_usage(true), 'memory limit reached');
        $this->assertLessThan($time_limit_seconds, (microtime(true) - $start_timestamp), 'time limit reached');
    }
    
    private function makeStream($filename)
    {
        return fopen($filename, 'w');
    }
    
    private function makeBookWithWriter($fp)
    {
        $book = $this->factory->getBook();
        $writer = $this->factory->getWriterFactory()->getCsvStreamWriter($fp);
        $book->setWriter($writer);
        return $book;
    }
}