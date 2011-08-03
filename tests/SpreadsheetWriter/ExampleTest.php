<?php
namespace SpreadSheetWriter;

require_once(dirname(__DIR__) . '/bootstrap.php');

use SpreadSheetWriter\Parser\DOM\DOMFactory;

class ExampleTest extends \PHPUnit_Framework_TestCase
{
    public function testIntegration()
    {
        $out_filename = __DIR__ . '/_files/out.xml';
        
        $factory = new DOMFactory();
        $book = $factory->getBook();
        $fp = fopen($out_filename, 'w');
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
        $sheet->addRow($factory->getRow(array('head1', 'head2', 'head3', 'head4'))
                               ->setStyle($style));
        
        for($i = 0; $i < 100; $i++) {
            $sheet->addRow($factory->getRow(array(
                'cell1',
                'cell2',
                'cell3',
                'cell4'
            )));
        }

        $book->close();
        
//        $this->assertFileEquals($expected_filename, $out_filename);
        $this->assertTrue(memory_get_peak_usage(true) < 10000000, 'Memory usage over limit: ' . memory_get_peak_usage(true));
        
        $actual = file_get_contents($out_filename);
        $this->assertSelectCount('Workbook Styles Style', 1, $actual);
        $this->assertSelectCount('Workbook Worksheet', 2, $actual);
        $this->assertSelectCount('Workbook Worksheet Table Row', 201, $actual);
    }
}