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

use SpreadSheetWriter\Row;
use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;
use SpreadSheetWriter\Style;
use SpreadSheetWriter\Writer\OfficeOpenXML2007\WriterBase;
use SpreadSheetWriter\Writer\OfficeOpenXML2007\SharedStringsHelper;

final class OfficeOpenXML2007StreamWriter extends WriterBase
{
    const FONT_FAMILY_DEFAULT = 'Arial';
    const FONT_SIZE_DEFAULT = 10;
    
    private $sheetStream;
    private $rowId = 0;
    
    /**
     * @var Style
     */
    private $defaultStyle;

    /**
     * @var SharedStringsHelper
     */
    private $sharedStrings;
    
    public function startBook(Book $book)
    {
        $this->createBaseStructure();
        $this->defaultStyle = $book->newStyle();
        $this->startSharedStrings();
    }

    private function startSharedStrings()
    {
        $sharedStringsFile = $this->dataDir . DIRECTORY_SEPARATOR . 'sharedStrings.xml';
        $sharedStringsStream = $this->createWorkingStream($sharedStringsFile);
        $this->sharedStrings = new SharedStringsHelper($sharedStringsStream);
        $this->sharedStrings->start();
    }
    
    public function startSheet(Book $book, Sheet $sheet)
    {
        $sheetFile = $this->sheetDir . DIRECTORY_SEPARATOR . 'sheet' . $sheet->getId() . '.xml';
        $this->sheetStream = $this->createWorkingStream($sheetFile);
        fwrite($this->sheetStream, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' . self::EOL);
        fwrite($this->sheetStream, '    <sheetData>' . self::EOL);
        $this->rowId = 0;
    }
    
    public function writeRow(Row $row)
    {
        $columnId = 'A';
        $rowId = ++$this->rowId;
        $out = '        <row>' . self::EOL;
        foreach($row->getCells() as $cell) {
            $out .= '            <c r="' . $columnId . $rowId . '"';
            $out .= ' s="' . ($row->getStyle() ? $row->getStyle()->getId() : $this->defaultStyle->getId()) . '"';
            if(is_numeric($cell)) {
                $out .= '><v>' . $cell . '</v></c>' . self::EOL;
            } else {
                $sharedStringId = $this->sharedStrings->writeString($this->escape($cell));
                $out .= ' t="s"><v>' . $sharedStringId . '</v></c>' . self::EOL;
            }
            $columnId++;
        }
        
        fwrite($this->sheetStream, $out . '        </row>' . self::EOL);
    }
    
    public function endSheet(Book $book, Sheet $sheet)
    {
        fwrite($this->sheetStream, '    </sheetData>' . self::EOL);
        fwrite($this->sheetStream, '</worksheet>');
        fclose($this->sheetStream);
    }
    
    public function endBook(Book $book)
    {
        $this->sharedStrings->end();
        $this->createBookFile($book->getSheets());
        $this->createStylesFile($book->getStyles());
        $this->createDataRelationsFile($book->getSheets());
        $this->createContentTypesFile($book->getSheets());
        $this->zipWorkingFiles();
        $this->copyZipToStream();
//        $this->cleanWorkingFiles();
    }
    
    private function createDataRelationsFile(array $sheets)
    {
        $data = '<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' . self::EOL;
        
        /* @var Sheet $sheet */
        foreach($sheets as $sheet) {
            $data .= '    <Relationship Id="rId' . $sheet->getId() . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $sheet->getId() . '.xml"/>' . self::EOL;
        }
        
        $data .= '</Relationships>';
        
        $filename = $this->dataRelsDir . DIRECTORY_SEPARATOR . 'workbook.xml.rels';
        $this->createWorkingFile($filename, $data);
    }
    
    private function createContentTypesFile(array $sheets)
    {
        $data = '<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' . self::EOL;
    
        /* @var Sheet $sheet */
        foreach($sheets as $sheet) {
            $data .= '    <Override PartName="/xl/worksheets/sheet' . $sheet->getId() . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' . self::EOL;
        }
        
        $data .= '    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>';

        $contentTypesFile = $this->baseDir . DIRECTORY_SEPARATOR . '[Content_Types].xml';
        $this->createWorkingFile($contentTypesFile, $data);
    }
    
    private function createBookFile(array $sheets)
    {
        $data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>' . self::EOL;
        
        /* @var Sheet $sheet */
        foreach($sheets as $sheet) {
            $data .= '        <sheet name="' . $this->escape($sheet->getName()) . '" sheetId="' . $sheet->getId() . '" r:id="rId' . $sheet->getId() . '"/>' . self::EOL;
        }
        
        $data .= '    </sheets>
</workbook>';
        
        $filename = $this->dataDir . DIRECTORY_SEPARATOR . 'workbook.xml';
        $this->createWorkingFile($filename, $data);
    }
    
    /**
     *
     * @var Style $style
     * @param Style[] $styles 
     */
    private function createStylesFile(array $styles)
    {
        $data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' . self::EOL;
        $data .= $this->buildStyleFonts($styles);
        $data .= $this->buildStyleCellXfs($styles);
        $data .= '</styleSheet>';
        
        $filename = $this->dataDir . DIRECTORY_SEPARATOR . 'styles.xml';
        $this->createWorkingFile($filename, $data);
    }
    
    private function buildStyleFonts(array $styles)
    {
        $data = '    <fonts count="' . count($styles) . '">' . self::EOL;
        foreach($styles as $style) {
            $data .= '        <font>' . self::EOL;
            if($style->getFontBold()) {
                $data .= '            <b/>' . self::EOL;
            }
            $data .= '            <sz val="' . ($style->getFontSize() ? $style->getFontSize() : self::FONT_SIZE_DEFAULT) . '"/>' . self::EOL;
            $data .= '            <name val="' . ($style->getFontFamily() ? $style->getFontFamily() : self::FONT_FAMILY_DEFAULT) . '"/>' . self::EOL;
            $data .= '            <family val="2"/>' . self::EOL; // no clue why this needs to be there
            $data .= '        </font>' . self::EOL;
        }
        $data .= '    </fonts>' . self::EOL;
        return $data;
    }
    
    private function buildStyleCellXfs(array $styles)
    {
        $i = 0;
        $data = '    <cellXfs count="' . count($styles) . '">' . self::EOL;
        foreach($styles as $style) {
            $data .= '        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' . self::EOL;
            $i++;
        }
        $data .= '    </cellXfs>' . self::EOL;
        return $data;
    }
}