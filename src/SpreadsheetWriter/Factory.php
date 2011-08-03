<?php
namespace SpreadSheetWriter;

interface Factory 
{
    public function getWriterFactory();
    
    public function getBook();
    
    public function getSheet();
    
    public function getRow(array $cells);
    
    public function getStyle($id);
}