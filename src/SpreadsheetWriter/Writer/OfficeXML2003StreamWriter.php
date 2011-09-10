<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\Writer;
use SpreadSheetWriter\Row;
use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;

class OfficeXml2003StreamWriter implements Writer
{
    const EOL = "\r\n";
    
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
        $this->writeStream('<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" 
          xmlns:c="urn:schemas-microsoft-com:office:component:spreadsheet" 
          xmlns:html="http://www.w3.org/TR/REC-html40" 
          xmlns:o="urn:schemas-microsoft-com:office:office" 
          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" 
          xmlns:x2="http://schemas.microsoft.com/office/excel/2003/xml" 
          xmlns:x="urn:schemas-microsoft-com:office:excel" 
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' . self::EOL);
    }
    
    public function endBook(Book $book)
    {
        $this->writeStream('</Workbook>');
    }
    
    public function startSheet(Sheet $sheet)
    {
        $this->writeStyles($sheet->getStyles());
        $this->writeStream('    <Worksheet ss:Name="' . $sheet->getName() . '">
        <Table>' . self::EOL);
    }
    
    private function writeStyles(array $styles)
    {
        if(! count($styles)) {
            return;
        }
        
        $out = '    <Styles>' . self::EOL;
        $out .= '        <Style ss:ID="Default" ss:Name="Default"/>' . self::EOL;
        /* @var $style Style */
        foreach($styles as $style) {
            $out .= '        <Style ss:ID="' . $style->getId() . '" ss:Name="' . $style->getId() . '">' . self::EOL;
            $out .= '            <Font';
            if($style->getFontFamily()) {
                $out .= ' x:Family="' . $style->getFontFamily() . '"';
            }
            if($style->getFontBold()) {
                $out .= ' ss:Bold="1"';
            }
            $out .= '/>' . self::EOL;
            $out .= '        </Style>' . self::EOL;
        }
        $out .= '    </Styles>' . self::EOL;
        $this->writeStream($out);
    }
    
    public function endSheet(Sheet $sheet)
    {
        $this->writeStream('        </Table>
        <x:WorksheetOptions/>
    </Worksheet>' . self::EOL);
    }
    
    public function writeRow(Row $row)
    {
        $strStyle = ($row->getStyle() ? ' ss:StyleID="' . $row->getStyle()->getId() . '"' : '');
        
        $out = '            <Row>';
        foreach($row->getCells() as $cell) {
            $out .= '<Cell' . $strStyle . '><Data ss:Type="' . (is_numeric($cell) ? 'Number' : 'String') . '">' . $cell . '</Data></Cell>';
        }
        
        $this->writeStream($out . '</Row>' . self::EOL);
    }
    
    private function writeStream($data)
    {
        fwrite($this->stream, $data);
    }
}