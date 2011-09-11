<?php
namespace SpreadSheetWriter\Writer\OfficeOpenXML2007;

use SpreadSheetWriter\Writer\OfficeOpenXML2007StreamWriter as MyWriter;
use SpreadSheetWriter\IOException;

final class SharedStringsHelper
{
    private $filename;
    private $id = 0;
    private $stream;
    private $headerInsertPosition;
    
    public function __construct($filename)
    {
        $this->filename = $filename;
    }
    
    public function start()
    {
        $this->stream = fopen($this->filename, 'w');
        if(! $this->stream) {
            throw new IOException('failed to open stream: ' . $this->filename);
        }

        // NOTE: we leave extra space, so we can fseek and put in the correct count and uniqueCount later
        $firstPartOfHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . MyWriter::EOL . '<sst';
        $this->headerInsertPosition = strlen($firstPartOfHeader);
        fwrite($this->stream, $firstPartOfHeader);
        fwrite($this->stream, ' count="9999999" uniqueCount="9999999" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' . MyWriter::EOL);
    }
    
    /**
     * String MUST already be escaped
     * @param string $string
     * @return integer id referencing string
     */
    public function writeString($string)
    {
        fwrite($this->stream, '    <si><t>' . Utils::escape($string) . '</t></si>' . MyWriter::EOL);
        return $this->id++;
    }
    
    public function end()
    {
        $stringCount = $this->id;
        fwrite($this->stream, '</sst>');
        fseek($this->stream, $this->headerInsertPosition);
        fwrite($this->stream, sprintf("%-38s", ' count="' . $stringCount . '" uniqueCount="' . $stringCount . '"'));
        fclose($this->stream);
    }
}