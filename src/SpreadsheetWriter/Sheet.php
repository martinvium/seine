<?php
namespace SpreadSheetWriter;

interface Sheet
{
    public function addRow(Row $row);
    
    public function getName();
    
    public function getStyles();
    
    /**
     * Sheet unique numeric id
     */
    public function getId();
    
    /**
     * @internal assigned by Book
     */
    public function setId($id);
}