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
namespace Seine\Writer;

use Seine\Writer;
use Seine\Row;
use Seine\Book;
use Seine\Sheet;
use Seine\Configuration;

/**
 * @link http://tools.ietf.org/html/rfc4180
 */
class CsvStreamWriter implements Writer
{
    const OPT_FIELD_DELIMITER = 'CsvStreamWriter::OPT_FIELD_DELIMITER';
    const OPT_TEXT_DELIMITER = 'CsvStreamWriter::OPT_TEXT_DELIMITER';
    const OPT_ROW_DELIMITER = 'CsvStreamWriter::OPT_ROW_DELIMITER';

    const CRLF = "\r\n";

    const FIELD_DEL_DEFAULT = ",";
    const TEXT_DEL_DEFAULT = '"';

    private $rowDelimiter = self::CRLF;
    private $fieldDelimiter = self::FIELD_DEL_DEFAULT;
    private $textDelimiter = self::TEXT_DEL_DEFAULT;
    private $autoCloseStream = false;

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

    public function setAutoCloseStream($flag) 
    {
        $this->autoCloseStream = (bool)$flag;
    }

    public function startBook(Book $book)
    {

    }

    public function endBook(Book $book)
    {
        if($this->autoCloseStream) {
            fclose($this->stream);
        }
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

    public function setConfig(Configuration $config)
    {
        $this->setFieldDelimiter($config->getOption(self::OPT_FIELD_DELIMITER, $this->fieldDelimiter));
        $this->setTextDelimiter($config->getOption(self::OPT_TEXT_DELIMITER, $this->textDelimiter));
        $this->setRowDelimiter($config->getOption(self::OPT_ROW_DELIMITER, $this->rowDelimiter));
        $this->setAutoCloseStream($config->getOption(Configuration::OPT_AUTO_CLOSE_STREAM, $this->autoCloseStream));
    }
}