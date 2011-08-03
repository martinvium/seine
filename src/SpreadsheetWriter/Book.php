<?php
namespace SpreadSheetWriter;

interface Book
{
    public function addSheet(Sheet $sheet);
}