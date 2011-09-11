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
namespace SpreadSheetWriter\Parser\DOM;

use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;
use SpreadSheetWriter\Writer;

final class DOMBook extends DOMElement implements Book
{
    private $sheetId = 1;
    private $sheets = array();
    
    /**
     * @var Sheet
     */
    private $lastSheet = null;
    
    /**
     * @var Writer
     */
    private $writer;
   
    private $started = false;
    
    private $styles = array();
    
    public function setWriter(Writer $writer)
    {
        $this->writer = $writer;
    }
    
    /**
     * @param string $name
     * @return DOMSheet
     */
    public function addSheetByName($name)
    {
        $sheet = $this->factory->getSheet();
        $sheet->setName($name);
        $this->addSheet($sheet);
        return $sheet;
    }
    
    /**
     * @param DOMSheet $sheet
     * @return DOMSheet 
     */
    public function addSheet(Sheet $sheet)
    {
        $sheet->setBook($this);
        $sheet->setId($this->sheetId++);
        $this->startBook();
        
        if($this->lastSheet) {
            $this->lastSheet->close();
        }
        
        $sheet->setWriter($this->writer);
        
        return $this->lastSheet = $this->sheets[] = $sheet;
    }
    
    public function getSheets()
    {
        return $this->sheets;
    }
    
    /**
     * @return DOMStyle
     */
    public function newStyle()
    {
        $style = $this->factory->getStyle(count($this->styles));
        $this->styles[] = $style;
        return $style;
    }
    
    /**
     * @return DOMStyle[]
     */
    public function getStyles()
    {
        return $this->styles;
    }
    
    private function startBook()
    {
        if(! $this->writer) {
            throw new \Exception('writer is undefined');
        }
        
        if($this->started) {
            return;
        }
        
        $this->writer->startBook($this);
        $this->started = true;
    }
    
    public function close()
    {
        if($this->lastSheet) {
            $this->lastSheet->close();
        }
        
        if($this->started) {
            $this->writer->endBook($this);
        }
        
        $this->started = false;
    }
    
    public function __destruct()
    {
        $this->close();
    }
}