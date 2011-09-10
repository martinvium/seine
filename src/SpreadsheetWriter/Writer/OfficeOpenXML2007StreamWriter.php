<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\Writer;
use SpreadSheetWriter\Row;
use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;

final class OfficeOpenXML2007StreamWriter extends OfficeOpenXML2007Base
{
    private $workbookStream;
    private $stylesStream;
    
    private $sheetId = 0;
    private $sheetStream;
    
    public function startBook(Book $book)
    {
        $this->createBaseStructure();
        $this->startBookStream();
        $this->startStylesStream();
    }
    
    private function startBookStream()
    {
        $workbookFile = $this->dataDir . DIRECTORY_SEPARATOR . 'workbook.xml';
        $this->workbookStream = $this->createWorkingStream($workbookFile);
        
        fwrite($this->workbookStream, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="Calc"/>
    <workbookPr backupFile="false" showObjects="all" dateCompatibility="false"/>
    <workbookProtection/>
    <bookViews>
        <workbookView/>
    </bookViews>
    <sheets>' . self::EOL);
    }
    
    private function startStylesStream()
    {
        $stylesFile = $this->dataDir . DIRECTORY_SEPARATOR . 'styles.xml';
        $this->stylesStream = $this->createWorkingStream($stylesFile);
        
        fwrite($this->stylesStream, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' . self::EOL);
    }
    
    public function startSheet(Sheet $sheet)
    {
        $this->sheetId++;
        fwrite($this->workbookStream, '        <sheet name="' . $this->escape($sheet->getName()) . '" sheetId="' . $this->sheetId . '" state="visible" r:id="' . $this->escape($sheet->getName()) . '"/>' . self::EOL);
        $this->writeStyles($sheet->getStyles());
        
        $sheetFile = $this->sheetDir . DIRECTORY_SEPARATOR . 'sheet' . $this->sheetId . '.xml';
        $this->sheetStream = $this->createWorkingStream($sheetFile);
        fwrite($this->sheetStream, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' . self::EOL);
        fwrite($this->sheetStream, '<sheetData>' . self::EOL);
    }
    
    private function writeStyles(array $styles)
    {
        // TODO
    }
    
    public function writeRow(Row $row)
    {
        $out = '            <row>';
        foreach($row->getCells() as $cell) {
            $out .= '<c' . (! is_numeric($cell) ? ' t="s"' : '') . '>' . $this->escape($cell) . '</c>';
        }
        
        fwrite($this->sheetStream, $out . '</row>' . self::EOL);
    }
    
    public function endSheet(Sheet $sheet)
    {
        fwrite($this->sheetStream, '</sheetData>' . self::EOL);
        fwrite($this->sheetStream, '</worksheet>');
        fclose($this->sheetStream);
    }
    
    public function endBook(Book $book)
    {
        $this->endBookStream();
        $this->endStylesStream();
        $this->zipWorkingFiles();
        $this->copyZipToStream();
//        $this->cleanWorkingFiles();
    }
    
    private function endBookStream()
    {
        fwrite($this->workbookStream, '    </sheets>
    <calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>
</workbook>');
        fclose($this->workbookStream);
    }
    
    private function endStylesStream()
    {
        fwrite($this->stylesStream, '</styleSheet>');
        fclose($this->stylesStream);
    }
}