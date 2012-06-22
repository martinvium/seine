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

    /**
     * @var Style
     */
    private $defaultPercentStyle;

    /**
     * @var Style
     */
    private $defaultDateStyle;

    private $filename;
    private $stream;
    private $rowId = 0;

    public function __construct(Sheet $sheet, SharedStringsHelper $sharedStrings, Style $defaultStyle, $filename, Style $defaultPercentStyle, Style $defaultDateStyle)
    {
        $this->sheet = $sheet;
        $this->sharedStrings = $sharedStrings;
        $this->defaultStyle = $defaultStyle;
        $this->defaultPercentStyle = $defaultPercentStyle;
        $this->defaultDateStyle = $defaultDateStyle;
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

    public function generateRow(Row $row, $rowId)
    {
        $columnId = 'A';

        $out = '       <row>' . MyWriter::EOL;
        foreach($row->getCells() as $cell) {
            $style = $row->getStyle() ? $row->getStyle()->getId() : $this->defaultStyle->getId();

            $out .= '            <c r="' . $columnId . $rowId . '"';

            if(is_numeric($cell)) {
                $out .= ' s="' . $style . '"';
                $out .= '><v>' . $cell . '</v></c>' . MyWriter::EOL;
            } elseif ($cell instanceof \DateTime) {
                $out .= ' s="' . $this->defaultDateStyle->getId() . '"';
                $out .= ' t="n"><v>' . $cell->diff(new \DateTime('1899-12-30'))->format('%a') . '</v></c>' . MyWriter::EOL;
            } elseif(strlen($cell) > 0 && substr($cell, -1) === '%' && is_numeric(substr($cell, 0, strlen($cell) - 1))) { // Percent
                $out .= ' s="' . $this->defaultPercentStyle->getId() . '"';
                $out .= ' t="n"><v>' . substr($cell, 0, strlen($cell) - 1)/100 . '</v></c>' . MyWriter::EOL;
            } else {
                $out .= ' s="' . $style . '"';
                $sharedStringId = $this->sharedStrings->writeString($cell);
                $out .= ' t="s"><v>' . $sharedStringId . '</v></c>' . MyWriter::EOL;
            }
            $columnId++;
        }

        return $out . '        </row>';
    }

    public function writeRow(Row $row)
    {
        $rowId = ++$this->rowId;

        fwrite($this->stream, $this->generateRow($row, $rowId) . MyWriter::EOL);
    }

    public function end()
    {
        fwrite($this->stream, '    </sheetData>' . MyWriter::EOL);
        fwrite($this->stream, '</worksheet>');
        fclose($this->stream);
    }
}