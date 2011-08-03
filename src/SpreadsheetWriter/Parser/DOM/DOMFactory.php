<?php
namespace SpreadSheetWriter\Parser\DOM;

use SpreadSheetWriter\Factory;
use SpreadSheetWriter\Writer\WriterFactoryImpl;

final class DOMFactory implements Factory 
{
    /**
     * @return WriterFactoryImpl 
     */
    public function getWriterFactory()
    {
        return new WriterFactoryImpl($this);
    }
    
    /**
     * @return DOMBook 
     */
    public function getBook()
    {
        return new DOMBook($this);
    }
    
    /**
     * @param array $cells
     * @return DOMArrayRow
     */
    public function getRow(array $cells)
    {
        return new DOMArrayRow($this, $cells);
    }
    
    /**
     * @return DOMSheet 
     */
    public function getSheet()
    {
        return new DOMSheet($this);
    }

    /**
     * @param string $id
     * @return DOMStyle
     */
    public function getStyle($id)
    {
        return new DOMStyle($this, $id);
    }
}