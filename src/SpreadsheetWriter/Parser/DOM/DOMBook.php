<?php
namespace SpreadSheetWriter\Parser\DOM;

use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;
use SpreadSheetWriter\Writer;

final class DOMBook extends DOMElement implements Book
{
    private $sheets = array();
    
    /**
     * @var Sheet
     */
    private $lastSheet = null;
    
    /**
     * @var Writer
     */
    private $writer;
   
    private $started = false;
    
    public function setWriter(Writer $writer)
    {
        $this->writer = $writer;
    }
    
    /**
     * @param string $name
     * @return DOMSheet
     */
    public function addSheetByName($name)
    {
        $sheet = $this->factory->getSheet();
        $sheet->setName($name);
        $this->addSheet($sheet);
        return $sheet;
    }
    
    /**
     * @param Sheet $sheet
     * @return DOMSheet 
     */
    public function addSheet(Sheet $sheet)
    {
        $sheet->setId(count($this->sheets));
        $this->startBook();
        
        if($this->lastSheet) {
            $this->lastSheet->close();
        }
        
        $sheet->setWriter($this->writer);
        
        return $this->lastSheet = $this->sheets[] = $sheet;
    }
    
    public function getSheets()
    {
        return $this->sheets;
    }
    
    private function startBook()
    {
        if(! $this->writer) {
            throw new \Exception('writer is undefined');
        }
        
        if($this->started) {
            return;
        }
        
        $this->writer->startBook($this);
        $this->started = true;
    }
    
    public function close()
    {
        if($this->lastSheet) {
            $this->lastSheet->close();
        }
        
        if($this->started) {
            $this->writer->endBook($this);
        }
        
        $this->started = false;
    }
    
    public function __destruct()
    {
        $this->close();
    }
}