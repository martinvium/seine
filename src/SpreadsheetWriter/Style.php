<?php
namespace SpreadSheetWriter;

interface Style 
{
    public function getId();
    
    public function getFontBold();
    
    public function getFontFamily();
    
    public function getFontSize();
}