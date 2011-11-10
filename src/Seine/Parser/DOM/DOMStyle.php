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
namespace Seine\Parser\DOM;

use Seine\Style;
use Seine\Factory;

final class DOMStyle implements Style
{
    private $factory;
    private $id = '';
    private $fontBold = false;
    private $fontFamily = '';
    private $fontSize = 0;
    
    public function __construct(Factory $factory, $id)
    {
        $this->factory = $factory;
        $this->id = (int)$id;
    }

    public function getId()
    {
        return $this->id;
    }
    
    /**
     * @param boolean $isBold
     * @return DOMStyle 
     */
    public function setFontBold($isBold)
    {
        $this->fontBold = (bool)$isBold;
        return $this;
    }

    /**
     * @param string $family
     * @return DOMStyle 
     */
    public function setFontFamily($family)
    {
        $this->fontFamily = (string)$family;
        return $this;
    }

    /**
     * @param integer $size
     * @return DOMStyle 
     */
    public function setFontSize($size)
    {
        $this->fontSize = (int)$size;
        return $this;
    }

    public function getFontBold()
    {
        return $this->fontBold;
    }

    public function getFontFamily()
    {
        return $this->fontFamily;
    }

    public function getFontSize()
    {
        return $this->fontSize;
    }
    
    
}