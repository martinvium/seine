<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\Writer;
use SpreadSheetWriter\Row;
use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;

/**
 * @link http://tools.ietf.org/html/rfc4180
 */
class CsvStreamWriter implements Writer
{
    const CRLF = "\r\n";
    
    const FIELD_DEL_DEFAULT = ",";
    const TEXT_DEL_DEFAULT = '"';
    
    private $rowDelimiter = self::CRLF;
    private $fieldDelimiter = self::FIELD_DEL_DEFAULT;
    private $textDelimiter = self::TEXT_DEL_DEFAULT;
    
    private $stream;
    
    public function __construct($stream)
    {
        if(! is_resource($stream)) {
            throw new \InvalidArgumentException('fp is not a valid stream resource');
        }
        
        $this->stream = $stream;
    }
    
    public function setRowDelimiter($del)
    {
        $this->rowDelimiter = $del;
    }
    
    public function setFieldDelimiter($del)
    {
        $this->fieldDelimiter = $del;
    }
    
    public function setTextDelimiter($del)
    {
        $this->textDelimiter = $del;
    }
    
    public function startBook(Book $book)
    {
        
    }
    
    public function endBook(Book $book)
    {
        
    }
    
    public function startSheet(Sheet $sheet)
    {
        
    }
    
    public function endSheet(Sheet $sheet)
    {
        
    }
    
    public function writeRow(Row $row)
    {
        $paddedCells = array_map(array($this, 'quote'), $row->getCells());
        $this->writeStream(implode($this->fieldDelimiter, $paddedCells) . $this->rowDelimiter);
    }
    
    private function writeStream($data)
    {
        fwrite($this->stream, $data);
    }
    
    private function quote($string)
    {
        $escapedString = str_replace($this->textDelimiter, $this->textDelimiter . $this->textDelimiter, $string);
        return $this->textDelimiter . $escapedString . $this->textDelimiter;
    }
}