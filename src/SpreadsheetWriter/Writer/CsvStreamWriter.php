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

/**
 * @link http://tools.ietf.org/html/rfc4180
 */
class CsvStreamWriter implements Writer
{
    const CRLF = "\r\n";

    const FIELD_DEL_DEFAULT = ",";
    const TEXT_DEL_DEFAULT = '"';

    private $rowDelimiter = self::CRLF;
    private $fieldDelimiter = self::FIELD_DEL_DEFAULT;
    private $textDelimiter = self::TEXT_DEL_DEFAULT;

    private $stream;

    public function __construct($stream)
    {
        if(! is_resource($stream)) {
            throw new \InvalidArgumentException('fp is not a valid stream resource');
        }

        $this->stream = $stream;
    }

    public function setRowDelimiter($del)
    {
        $this->rowDelimiter = $del;
    }

    public function setFieldDelimiter($del)
    {
        $this->fieldDelimiter = $del;
    }

    public function setTextDelimiter($del)
    {
        $this->textDelimiter = $del;
    }

    public function startBook(Book $book)
    {

    }

    public function endBook(Book $book)
    {

    }

    public function startSheet(Book $book, Sheet $sheet)
    {

    }

    public function endSheet(Book $book, Sheet $sheet)
    {

    }

    public function writeRow(Sheet $sheet, Row $row)
    {
        $paddedCells = array_map(array($this, 'quote'), $row->getCells());
        $this->writeStream(implode($this->fieldDelimiter, $paddedCells) . $this->rowDelimiter);
    }

    private function writeStream($data)
    {
        fwrite($this->stream, $data);
    }

    private function quote($string)
    {
        $escapedString = str_replace($this->textDelimiter, $this->textDelimiter . $this->textDelimiter, $string);
        return $this->textDelimiter . $escapedString . $this->textDelimiter;
    }
}