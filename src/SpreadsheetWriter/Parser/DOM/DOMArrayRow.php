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

use SpreadSheetWriter\Row;
use SpreadSheetWriter\Factory;
use SpreadSheetWriter\Style;

final class DOMArrayRow extends DOMElement implements Row
{
    private $cells = array();
    
    /**
     * @var Style
     */
    private $style;
    
    public function __construct(Factory $factory, array $cells)
    {
        parent::__construct($factory);
        $this->cells = $cells;
    }
    
    public function getCells()
    {
        return $this->cells;
    }
    
    /**
     * @param array $cells
     * @return DOMArrayRow 
     */
    public function setCells(array $cells)
    {
        $this->cells = $cells;
        return $this;
    }

    public function getStyle()
    {
        return $this->style;
    }
    
    /**
     * @param Style $style
     * @return DOMArrayRow 
     */
    public function setStyle(Style $style)
    {
        $this->style = $style;
        return $this;
    }
}