<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\WriterFactory;

class WriterFactoryImpl implements WriterFactory
{
    /**
     * @param stream $stream
     * @return OfficeOpenXML2007StreamWriter 
     */
    public function getOfficeOpenXML2007StreamWriter($stream)
    {
        return new OfficeOpenXML2007StreamWriter($stream);
    }
    
    /**
     * @param stream $stream
     * @return OfficeXml2003StreamWriter 
     */
    public function getOfficeXML2003StreamWriter($stream)
    {
        return new OfficeXml2003StreamWriter($stream);
    }
    
    /**
     * @param stream $stream
     * @return CsvStreamWriter 
     */
    public function getCsvStreamWriter($stream)
    {
        return new CsvStreamWriter($stream);
    }
}