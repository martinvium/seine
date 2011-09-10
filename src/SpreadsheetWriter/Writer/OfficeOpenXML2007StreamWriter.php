<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\Row;
use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;

final class OfficeOpenXML2007StreamWriter extends OfficeOpenXML2007Base
{
    private $sheetStream;
    
    private $rowId = 0;
    
    public function startBook(Book $book)
    {
        $this->createBaseStructure();
    }
    
    public function startSheet(Sheet $sheet)
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
            if(is_numeric($cell)) {
                $out .= '><v>' . $cell . '</v></c>' . self::EOL;
            } else {
                $out .= ' t="s"><v>' . $this->escape($cell) . '</v></c>' . self::EOL;
            }
            $columnId++;
        }
        
        fwrite($this->sheetStream, $out . '        </row>' . self::EOL);
    }
    
    public function endSheet(Sheet $sheet)
    {
        fwrite($this->sheetStream, '    </sheetData>' . self::EOL);
        fwrite($this->sheetStream, '</worksheet>');
        fclose($this->sheetStream);
    }
    
    public function endBook(Book $book)
    {
        $this->createBookFile($book->getSheets());
        $this->createStylesFile($book->getSheets());
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
        
        $data .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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
    
    private function createStylesFile()
    {
        $data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' . self::EOL;
        $data .= '</styleSheet>';
        
        $filename = $this->dataDir . DIRECTORY_SEPARATOR . 'styles.xml';
        $this->createWorkingFile($filename, $data);
    }
}