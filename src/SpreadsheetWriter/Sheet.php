<?php
namespace SpreadSheetWriter;

interface Sheet
{
    public function addRow(Row $row);
    
    public function getName();
    
    public function getStyles();
}