<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\Writer;
use SpreadSheetWriter\Row;
use SpreadSheetWriter\Book;
use SpreadSheetWriter\Sheet;

class CsvStreamWriter implements Writer
{
    const ROW_DEL_WINDOWS = "\r\n";
    const ROW_DEL_MACOS = "\r";
    const ROW_DEL_UNIX = "\n";
    
    const COLUMN_DEL_DEFAULT = ";";
    
    private $rowDelimiter = self::ROW_DEL_WINDOWS;
    private $fieldDelimiter = self::COLUMN_DEL_DEFAULT;
    private $textDelimiter = '"';
    
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
        $paddedCells = array_map(array($this, 'padCell'), $row->getCells());
        $this->writeStream(implode($this->fieldDelimiter, $paddedCells) . $this->rowDelimiter);
    }
    
    private function padCell($val)
    {
        return $this->textDelimiter . $val . $this->textDelimiter;
    }
    
    private function writeStream($data)
    {
        fwrite($this->stream, $data);
    }
}