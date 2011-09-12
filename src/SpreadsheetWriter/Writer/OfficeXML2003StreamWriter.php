<?php
/**
 * Copyright (C) 2011 by Martin Vium
 * Copyright (C) 2011 by Kim Hemsø Rasmussen
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
use SpreadSheetWriter\Writer\OfficeXml2003StreamWriter\Sheet as Xml2003Sheet;

class OfficeXml2003StreamWriter implements Writer
{
    const CHARSET = 'utf-8';
    const EOL = "\r\n";

    private $stream;

    private $sheets = array();

    public function __construct($stream)
    {
        if(! is_resource($stream)) {
            throw new \InvalidArgumentException('fp is not a valid stream resource');
        }

        $this->stream = $stream;
    }

    public function startBook(Book $book)
    {
        $this->writeStream($this->stream, '<?xml version="1.0" encoding="' . self::CHARSET . '"?>' . self::EOL);
        $this->writeStream($this->stream, '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
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
        $this->writeStyles($book->getStyles());
        $this->writeSheets($this->sheets);
        $this->writeStream($this->stream, '</Workbook>');
    }

    public function startSheet(Book $book, Sheet $sheet)
    {
        $stream = fopen('php://temp', 'w+');
        $dataSheet = new Xml2003Sheet($stream);
        $dataSheet->setOpen(true);
        $this->sheets[$sheet->getId()] = $dataSheet;
        $out = '    <Worksheet ss:Name="' . $this->escape($sheet->getName()) . '"><Table>' . self::EOL;
        $this->writeStream($dataSheet->getStream(), $out);
    }

    /**
     * Write $sheets to output stream and closes each sheet's stream.
     *
     * @param array $sheets
     */
    private function writeSheets($sheets) {
        foreach ($sheets as $sheet) {
            fseek($sheet->getStream(), 0);
            stream_copy_to_stream($sheet->getStream(), $this->stream);
            fclose($sheet->getStream());
        }
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
        $this->writeStream($this->stream, $out);
    }

    public function endSheet(Book $book, Sheet $sheet)
    {
        $out = <<<EOD
        </Table>
        <x:WorksheetOptions/>
    </Worksheet>'
EOD;
        $dataSheet = $this->sheets[$sheet->getId()];
        $this->writeStream($dataSheet->getStream(), $out);
        $dataSheet->setOpen(false);
    }

    public function writeRow(Sheet $sheet, Row $row)
    {
        $dataSheet = $this->sheets[$sheet->getId()];

        if (!$dataSheet->isOpen()) {
            throw new \Exception(sprintf("Cant write row: sheet already closed (sheet: '%s')", $sheet->getId()));
        }

        $strStyle = ($row->getStyle() ? ' ss:StyleID="' . $this->escape($row->getStyle()->getId()) . '"' : '');

        $out = '            <Row>';
        foreach($row->getCells() as $cell) {
            $out .= '<Cell' . $strStyle . '><Data ss:Type="' . (is_numeric($cell) ? 'Number' : 'String') . '">' . $this->escape($cell) . '</Data></Cell>';
        }
        $out .= '</Row>' . self::EOL;

        $this->writeStream($dataSheet->getStream(), $out);
    }

    private function escape($string)
    {
        return htmlspecialchars($string, ENT_QUOTES, 'utf-8');
    }

    private function writeStream($stream, $data)
    {
        fwrite($stream, $data);
    }

    public function setConfig(Configuration $config)
    {

    }
}