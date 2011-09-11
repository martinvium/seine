<?php
namespace SpreadSheetWriter\Writer\OfficeOpenXML2007;

use SpreadSheetWriter\Row;
use SpreadSheetWriter\Sheet;
use SpreadSheetWriter\Style;
use SpreadSheetWriter\Writer\OfficeOpenXML2007StreamWriter as MyWriter;
use SpreadSheetWriter\IOException;

final class SheetHelper
{
    /**
     * @var Sheet
     */
    private $sheet;

    /**
     * @var SharedStringsHelper
     */
    private $sharedStrings;

    /**
     * @var Style
     */
    private $defaultStyle;

    private $filename;
    private $stream;
    private $rowId = 0;

    public function __construct(Sheet $sheet, SharedStringsHelper $sharedStrings, Style $defaultStyle, $filename)
    {
        $this->sheet = $sheet;
        $this->sharedStrings = $sharedStrings;
        $this->defaultStyle = $defaultStyle;
        $this->filename = $filename;
    }

    public function start()
    {
        $this->stream = fopen($this->filename, 'w');
        if(! $this->stream) {
            throw new IOException('failed to open stream: ' . $this->filename);
        }

        fwrite($this->stream, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' . MyWriter::EOL);
        fwrite($this->stream, '    <sheetData>' . MyWriter::EOL);
    }

    public function writeRow(Row $row)
    {
        $columnId = 'A';
        $rowId = ++$this->rowId;
        $out = '        <row>' . MyWriter::EOL;
        foreach($row->getCells() as $cell) {
            $out .= '            <c r="' . $columnId . $rowId . '"';
            $out .= ' s="' . ($row->getStyle() ? $row->getStyle()->getId() : $this->defaultStyle->getId()) . '"';
            if(is_numeric($cell)) {
                $out .= '><v>' . $cell . '</v></c>' . MyWriter::EOL;
            } else {
                $sharedStringId = $this->sharedStrings->writeString($cell);
                $out .= ' t="s"><v>' . $sharedStringId . '</v></c>' . MyWriter::EOL;
            }
            $columnId++;
        }

        fwrite($this->stream, $out . '        </row>' . MyWriter::EOL);
    }

    public function end()
    {
        fwrite($this->stream, '    </sheetData>' . MyWriter::EOL);
        fwrite($this->stream, '</worksheet>');
        fclose($this->stream);
    }
}