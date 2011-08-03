<?php
namespace SpreadSheetWriter\Parser\DOM;

use SpreadSheetWriter\Style;
use SpreadSheetWriter\Factory;

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
        $this->id = $id;
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