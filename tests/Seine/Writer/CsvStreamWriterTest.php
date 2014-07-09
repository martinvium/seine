<?php
/**
 * Copyright (C) 2011 by Martin Vium
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */
namespace Seine\Writer;

require_once(dirname(dirname(__DIR__)) . '/bootstrap.php');

use Seine\Parser\DOM\DOMFactory;
use Seine\Configuration;

class CsvStreamWriterTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var DOMFactory
     */
    private $factory;

    /**
     * @var Configuration
     */
    private $config;

    private $seine;

    public function setUp()
    {
        parent::setUp();
        $this->config = new Configuration();
        $this->config->setOption(Configuration::OPT_WRITER, 'CSVStreamWriter');
        $this->factory = new DOMFactory();
        $this->seine = new \Seine\Seine(array('writer' => 'csv'));
    }

    public function testWritesValidFormat()
    {
        $actual_file = __DIR__ . '/_tmp/actual_valid.csv';

        $doc = $this->seine->newDocument($actual_file);
        $sheet = $doc->newSheet('more1');
        for($i = 0; $i < 10; $i++) {
            $sheet->addRow(array(
                'cell1',
                'cell"2',
                'cæøå3',
                'cell4'
            ));
        }
        $doc->close();

        $this->assertFileEquals(__DIR__ . '/_files/expected_valid.csv', $actual_file);
    }

    public function testCustomStream()
    {
        $actual_file = __DIR__ . '/_tmp/actual_valid.csv';
        $fp = fopen($actual_file, 'w');

        $doc = $this->seine->newDocumentFromStream($fp);
        $sheet = $doc->newSheet('more1');
        for($i = 0; $i < 10; $i++) {
            $sheet->addRow(array(
                'cell1',
                'cell"2',
                'cæøå3',
                'cell4'
            ));
        }
        $doc->close();
        fclose($fp);

        $this->assertFileEquals(__DIR__ . '/_files/expected_valid.csv', $actual_file);
    }

    public function testWriteAllowsCustomizingDelimiters()
    {
        $actual_file = __DIR__ . '/_tmp/actual_custom_delimiters.csv';

        $this->seine->setOption(CsvStreamWriter::OPT_FIELD_DELIMITER, ';');
        $this->seine->setOption(CsvStreamWriter::OPT_TEXT_DELIMITER, "'");
        $this->seine->setOption(CsvStreamWriter::OPT_ROW_DELIMITER, 'NEWROW');

        $doc = $this->seine->newDocument($actual_file);
        $sheet = $doc->newSheet('more1');
        for($i = 0; $i < 10; $i++) {
            $sheet->addRow(array(
                'cell1',
                'ce"2',
                'cæøå3',
                'ce\'4'
            ));
        }
        $doc->close();

        $this->assertFileEquals(__DIR__ . '/_files/expected_custom_delimiters.csv', $actual_file);
    }

    /**
     * @group performance
     */
    public function testWriteSpeedAndMemoryUsage()
    {
        $memory_limit = 20 * 1000 * 1000; // 20 MB
        $time_limit_seconds = 5.0; // 5 seconds
        $num_rows = 10000;
        $num_columns = 100;

        $start_timestamp = microtime(true);
        $actual_file = __DIR__ . '/_files/performance.csv';

        $doc = $this->seine->newDocument($actual_file);
        $sheet = $doc->newSheet('more1');
        for($i = 0; $i < $num_rows; $i++) {
            $sheet->addRow(range(0, $num_columns));
        }
        $doc->close();

        $this->assertLessThan($memory_limit, memory_get_peak_usage(true), 'memory limit reached');
        $this->assertLessThan($time_limit_seconds, (microtime(true) - $start_timestamp), 'time limit reached');
    }
}