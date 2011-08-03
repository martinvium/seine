<?php
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