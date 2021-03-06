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
namespace Seine\Tests\Writer;

require_once(dirname(dirname(__DIR__)) . '/bootstrap.php');

use Seine\Seine;

class OfficeXML2003StreamWriterTest extends \PHPUnit_Framework_TestCase
{
    private $seine;

    public function setUp()
    {
        parent::setUp();
        $this->seine = new Seine(array('writer' => 'oxml2003'));
    }

    public function testMultipleSheetsWithStyles()
    {
        $actual_file = __DIR__ . '/_tmp/actual_valid.xml';


        $doc = $this->seine->newDocument($actual_file);
        $sheet = $doc->newSheet('more1');
        for($i = 0; $i < 10; $i++) {
            $sheet->addRow(array(
                'cell1',
                'cell"2',
                'cell</Cell>3',
                'cell4'
            ));
        }

        $sheet = $doc->newSheet('mor"e2');
        $style = $doc->newStyle()->setFontBold(true);
        $sheet->addRow($this->seine->getRow(array('head1', 'head2', 'head3', 'head4'))->setStyle($style));
        for($i = 0; $i < 10; $i++) {
            $sheet->addRow(array(
                'cell1',
                'ceæøåll2',
                rand(100, 10000),
                'cell4'
            ));
        }
        $doc->close();

        $actual = file_get_contents($actual_file);
        $this->assertSelectCount('Workbook Styles Style', 2, $actual, 'default style + one custom style');
        $this->assertSelectCount('Workbook Worksheet', 2, $actual, 'total number of sheets');
        $this->assertSelectCount('Workbook Worksheet Table Row', 21, $actual, 'total number of rows');
    }

    /**
     * @group performance
     */
    public function testWriteSpeedAndMemoryUsage()
    {
        $this->markTestIncomplete('does not actualy write any rows...');
        $memory_limit = 20 * 1000 * 1000; // 20 MB
        $time_limit_seconds = 8.0; // 8 seconds
        $num_rows = 10000;
        $num_columns = 100;

        $start_timestamp = microtime(true);
        $actual_file = __DIR__ . '/_tmp/performance.xml';

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