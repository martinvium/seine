<?php
namespace SpreadSheetWriter;

interface WriterFactory 
{
    public function getOfficeOpenXML2007StreamWriter($stream);
    
    public function getOfficeXML2003StreamWriter($stream);
    
    public function getCsvStreamWriter($stream);
}