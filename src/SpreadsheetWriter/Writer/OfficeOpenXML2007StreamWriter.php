<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\Writer;
use SpreadSheetWriter\Row;
use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;

final class OfficeOpenXML2007StreamWriter extends OfficeOpenXML2007Base
{
    private $dataDir;
    private $workbookStream;
    private $sheetIds = 1;
    
    public function startBook(Book $book)
    {
        $this->createBaseDir();
        $this->createContentTypesFile();
        $this->createRelsFile();
        $this->createDocPropFiles();
        $this->createDataDir();
        $this->startBookStream();
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
    
    private function createDataDir()
    {
        $this->dataDir = $this->baseDir . DIRECTORY_SEPARATOR . 'xl';
        $this->createWorkingDir($this->dataDir);
    }
    
    public function startSheet(Sheet $sheet)
    {
        fwrite($this->workbookStream, '        <sheet name="' . $this->escape($sheet->getName()) . '" sheetId="' . $this->sheetIds++ . '" state="visible" r:id="' . $this->escape($sheet->getName()) . '"/>' . self::EOL);
    }
    
    public function endSheet(Sheet $sheet)
    {
        
    }
    
    public function endBook(Book $book)
    {
        $this->endBookStream();
        $this->zipWorkingFiles();
        $this->copyZipToStream();
        $this->cleanWorkingFiles();
    }
    
    private function endBookStream()
    {
        fwrite($this->workbookStream, '    </sheets>
    <calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>
</workbook>');
        fclose($this->workbookStream);
    }

    public function writeRow(Row $row)
    {
        
    }
}