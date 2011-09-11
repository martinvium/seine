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
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\Writer;
use SpreadSheetWriter\Row;
use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;

class OfficeXml2003StreamWriter implements Writer
{
    const CHARSET = 'utf-8';
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
        $this->writeStream('<?xml version="1.0" encoding="' . self::CHARSET . '"?>' . self::EOL);
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
        $this->writeStream('    <Worksheet ss:Name="' . $this->escape($sheet->getName()) . '">
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
            $out .= '        <Style ss:ID="' . $this->escape($style->getId()) . '" ss:Name="' . $this->escape($style->getId()) . '">' . self::EOL;
            $out .= '            <Font';
            if($style->getFontFamily()) {
                $out .= ' x:Family="' . $this->escape($style->getFontFamily()) . '"';
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
        $strStyle = ($row->getStyle() ? ' ss:StyleID="' . $this->escape($row->getStyle()->getId()) . '"' : '');
        
        $out = '            <Row>';
        foreach($row->getCells() as $cell) {
            $out .= '<Cell' . $strStyle . '><Data ss:Type="' . (is_numeric($cell) ? 'Number' : 'String') . '">' . $this->escape($cell) . '</Data></Cell>';
        }
        
        $this->writeStream($out . '</Row>' . self::EOL);
    }
    
    private function escape($string)
    {
        return htmlspecialchars($string, ENT_QUOTES, 'utf-8');
    }
    
    private function writeStream($data)
    {
        fwrite($this->stream, $data);
    }
}