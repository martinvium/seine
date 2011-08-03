<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\WriterFactory;

class WriterFactoryImpl implements WriterFactory
{
    public function getOfficeOpenXML2007StreamWriter($stream)
    {
        return new OfficeOpenXML2007StreamWriter($stream);
    }
    
    public function getOfficeXML2003StreamWriter($stream)
    {
        return new OfficeXml2003StreamWriter($stream);
    }
}