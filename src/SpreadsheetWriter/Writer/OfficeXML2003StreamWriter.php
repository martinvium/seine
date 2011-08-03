<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\Writer;
use SpreadSheetWriter\Row;
use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;

class OfficeXml2003StreamWriter implements Writer
{
    private $stream;
    
    public function __construct($stream)
    {
        if(! is_resource($stream)) {
            throw new \InvalidArgumentException('fp is not a valid stream resource');
        }
        
        $this->stream = $stream;
    }
    
    public function startBook(Book $book)
    {
        $this->writeStream('<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook
   xmlns="urn:schemas-microsoft-com:office:spreadsheet"
   xmlns:o="urn:schemas-microsoft-com:office:office"
   xmlns:x="urn:schemas-microsoft-com:office:excel"
   xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
   xmlns:html="http://www.w3.org/TR/REC-html40">' . PHP_EOL);
    }
    
    public function endBook(Book $book)
    {
        $this->writeStream('</Workbook>');
    }
    
    public function startSheet(Sheet $sheet)
    {
        $this->writeStyles($sheet->getStyles());
        $this->writeStream('    <Worksheet ss:Name="' . $sheet->getName() . '">
        <Table x:FullColumns="1" x:FullRows="1">' . PHP_EOL);
    }
    
    private function writeStyles(array $styles)
    {
        if(! count($styles)) {
            return;
        }
        
        $out = '    <Styles>' . PHP_EOL;
        
        /* @var $style Style */
        foreach($styles as $style) {
            $out .= '        <Style ss:ID="' . $style->getId() . '">' . PHP_EOL;
            $out .= '            <Font';
            if($style->getFontFamily()) {
                $out .= ' x:Family="' . $style->getFontFamily() . '"';
            }
            if($style->getFontBold()) {
                $out .= ' ss:Bold="1"';
            }
            $out .= '/>' . PHP_EOL;
            $out .= '        </Style>' . PHP_EOL;
        }
        $out .= '    </Styles>' . PHP_EOL;
        $this->writeStream($out);
    }
    
    public function endSheet(Sheet $sheet)
    {
        $this->writeStream('        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
            <Print>
                <ValidPrinterInfo />
                <HorizontalResolution>600</HorizontalResolution>
                <VerticalResolution>600</VerticalResolution>
            </Print>
            <Selected />
            <Panes>
                <Pane>
                    <Number>3</Number>
                    <ActiveRow>5</ActiveRow>
                    <ActiveCol>1</ActiveCol>
                </Pane>
            </Panes>
            <ProtectObjects>False</ProtectObjects>
            <ProtectScenarios>False</ProtectScenarios>
        </WorksheetOptions>
    </Worksheet>' . PHP_EOL);
    }
    
    public function writeRow(Row $row)
    {
        $out = '            <Row>';
        foreach($row->getCells() as $cell) {
            $out .= '<Cell';
            if($row->getStyle()) {
                $out .= ' ss:StyleID="' . $row->getStyle()->getId() . '"';
            }
            $out .= '>';
            
            $out .= '<Data ss:Type="' . $this->formatType($cell) . '">';
            $out .= $this->formatValue($cell);
            $out .= '</Data></Cell>';
        }
        $out .= '</Row>' . PHP_EOL;
        $this->writeStream($out);
    }
    
    private function formatType($cell)
    {
        if(is_numeric($cell)) {
            return 'Number';
        } else {
            return 'String';
        }
    }
    
    private function formatValue($cell)
    {
        return $cell;
    }
    
    private function writeStream($data)
    {
        fwrite($this->stream, $data);
    }
}