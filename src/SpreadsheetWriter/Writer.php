<?php
namespace SpreadSheetWriter;

use SpreadSheetWriter\Row;
use SpreadSheetWriter\Sheet;
use SpreadSheetWriter\Book;

interface Writer 
{
    public function writeRow(Row $row);
    
    public function startSheet(Sheet $sheet);
    
    public function endSheet(Sheet $sheet);
    
    public function startBook(Book $book);
    
    public function endBook(Book $book);
}