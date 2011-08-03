<?php
namespace SpreadSheetWriter;

interface Row
{
    public function getCells();
    
    /**
     * @return Style
     */
    public function getStyle();
}