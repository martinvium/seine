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
namespace Seine\Writer\OfficeOpenXML2007;

use Seine\Row;
use Seine\Sheet;
use Seine\Style;
use Seine\Writer\OfficeOpenXML2007StreamWriter as MyWriter;
use Seine\IOException;

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
        $out = '        <row r="' . $rowId . '">' . MyWriter::EOL;
        foreach($row->getCells() as $cell) {
            $out .= '            <c r="' . $columnId . $rowId . '"';
            $out .= ' s="' . ($row->getStyle() ? $row->getStyle()->getId() : $this->defaultStyle->getId()) . '"';
            if(is_numeric($cell)) {
                $out .= '><v>' . $cell . '</v></c>' . MyWriter::EOL;
            } else {
                if (empty($cell)) {
                    $out .= '/>' . MyWriter::EOL;
                } else {
                    $sharedStringId = $this->sharedStrings->writeString($cell);
                    $out .= ' t="s"><v>' . $sharedStringId . '</v></c>' . MyWriter::EOL;
                }
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