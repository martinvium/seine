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

use Seine\Writer\OfficeOpenXML2007StreamWriter as MyWriter;
use Seine\IOException;

final class SharedStringsHelper
{
    /**
     * This number must be really big so that the no generated file will have more strings than that.
     * If the strings number goes above, characters will be overwritten in an unwanted way and will corrupt the file.
     */
    const DEFAULT_STRINGS_COUNT_PART = 'count="9999999999999" uniqueCount="9999999999999"';

    private $filename;
    private $id = 0;
    private $stream;
    private $headerInsertPosition;
    
    public function __construct($filename)
    {
        $this->filename = $filename;
    }
    
    public function start()
    {
        $this->stream = fopen($this->filename, 'w');
        if(! $this->stream) {
            throw new IOException('failed to open stream: ' . $this->filename);
        }

        // NOTE: we leave extra space, so we can fseek and put in the correct count and uniqueCount later
        $firstPartOfHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . MyWriter::EOL . '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ';
        $this->headerInsertPosition = strlen($firstPartOfHeader);
        fwrite($this->stream, $firstPartOfHeader);
        fwrite($this->stream, self::DEFAULT_STRINGS_COUNT_PART . '>' . MyWriter::EOL);
    }
    
    /**
     * String MUST already be escaped
     * @param string $string
     * @return integer id referencing string
     */
    public function writeString($string)
    {
        fwrite($this->stream, '    <si><t>' . Utils::escape($string) . '</t></si>' . MyWriter::EOL);
        return $this->id++;
    }
    
    public function end()
    {
        $stringCount = $this->id;
        fwrite($this->stream, '</sst>');

        // Replace the default strings count with the actual number of shared strings in the file header
        $defaultStringsCountPartLength = strlen(self::DEFAULT_STRINGS_COUNT_PART);
        fseek($this->stream, $this->headerInsertPosition);
        fwrite($this->stream, sprintf("%-{$defaultStringsCountPartLength}s", 'count="' . $stringCount . '" uniqueCount="' . $stringCount . '"'));

        fclose($this->stream);
    }
}