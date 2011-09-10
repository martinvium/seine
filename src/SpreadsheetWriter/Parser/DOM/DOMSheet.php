<?php
namespace SpreadSheetWriter\Parser\DOM;

use SpreadSheetWriter\Sheet;
use SpreadSheetWriter\Row;
use SpreadSheetWriter\Writer;

final class DOMSheet extends DOMElement implements Sheet
{
    /**
     * @var Writer
     */
    private $writer;
    
    private $name = '';
    
    private $started = false;
    
    private $styles = array();
    
    private $id;
    
    public function getId()
    {
        return $this->id;
    }
    
    /**
     * @internal assigned by Book
     */
    public function setId($id)
    {
        $this->id = (int)$id;
    }
    
    /**
     * @internal we do not store anything but the last row
     * @param Row $row 
     */
    public function addRow(Row $row)
    {
        $this->startSheet();
        $this->writer->writeRow($row);
    }
    
    /**
     * @param string $id
     * @return DOMStyle
     */
    public function addStyleById($id)
    {
        $style = $this->factory->getStyle($id);
        $this->styles[] = $style;
        return $style;
    }
    
    public function getStyles()
    {
        return $this->styles;
    }
    
    public function setName($name)
    {
        $this->name = (string)$name;
    }
    
    public function getName()
    {
        return $this->name;
    }
    
    public function setWriter(Writer $writer)
    {
        $this->writer = $writer;
    }
    
    private function startSheet()
    {
        if(! $this->writer) {
            throw new \Exception('writer is undefined');
        }
        
        if($this->started) {
            return;
        }
        
        $this->writer->startSheet($this);
        $this->started = true;
    }
    
    public function close()
    {
        if($this->started) {
            $this->writer->endSheet($this);
        }
        
        $this->started = false;
    }
}