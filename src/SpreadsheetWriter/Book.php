<?php
namespace SpreadSheetWriter;

interface Book
{
    /**
     * @param Sheet $sheet
     */
    public function addSheet(Sheet $sheet);
    
    /**
     * @return Sheet[]
     */
    public function getSheets();
}