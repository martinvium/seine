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

use SpreadSheetWriter\Sheet;
use SpreadSheetWriter\Row;
use SpreadSheetWriter\Writer;

final class DOMSheet extends DOMElement implements Sheet
{
    /**
     * @var Writer
     */
    private $writer;
    
    private $name = '';
    
    private $started = false;
    
    private $styles = array();
    
    private $id;
    
    public function getId()
    {
        return $this->id;
    }
    
    /**
     * @internal assigned by Book
     */
    public function setId($id)
    {
        $this->id = (int)$id;
    }
    
    /**
     * @internal we do not store anything but the last row
     * @param Row $row 
     */
    public function addRow(Row $row)
    {
        $this->startSheet();
        $this->writer->writeRow($row);
    }
    
    /**
     * @param string $id
     * @return DOMStyle
     */
    public function addStyleById($id)
    {
        $style = $this->factory->getStyle($id);
        $this->styles[] = $style;
        return $style;
    }
    
    public function getStyles()
    {
        return $this->styles;
    }
    
    public function setName($name)
    {
        $this->name = (string)$name;
    }
    
    public function getName()
    {
        return $this->name;
    }
    
    public function setWriter(Writer $writer)
    {
        $this->writer = $writer;
    }
    
    private function startSheet()
    {
        if(! $this->writer) {
            throw new \Exception('writer is undefined');
        }
        
        if($this->started) {
            return;
        }
        
        $this->writer->startSheet($this);
        $this->started = true;
    }
    
    public function close()
    {
        if($this->started) {
            $this->writer->endSheet($this);
        }
        
        $this->started = false;
    }
}