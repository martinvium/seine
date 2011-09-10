<?php
namespace SpreadSheetWriter\Writer;

use SpreadSheetWriter\Writer;
use SpreadSheetWriter\IOException;

abstract class OfficeOpenXML2007Base implements Writer
{
    const APP_NAME = 'SpreadSheetWriter';
    const EOL = "\r\n";
    
    private $workingFiles = array();
    
    protected $stream;
    protected $tempDir;
    protected $baseDir;
    protected $dataDir;
    protected $sheetDir;
    
    public function __construct($stream)
    {
        if(! is_resource($stream)) {
            throw new \InvalidArgumentException('fp is not a valid stream resource');
        }
        
        $this->stream = $stream;
        $this->tempDir = sys_get_temp_dir();
    }
    
    public function setTempDir($dir)
    {
        $this->tempDir = $dir;
    }
    
    protected function createBaseStructure()
    {
        $this->createBaseDir();
        $this->createContentTypesFile();
        $this->createRelsFile();
        $this->createDocPropFiles();
        $this->createDataDir();
        $this->createSheetDir();
    }
    
    private function createDataDir()
    {
        $this->dataDir = $this->baseDir . DIRECTORY_SEPARATOR . 'xl';
        $this->createWorkingDir($this->dataDir);
    }
    
    private function createSheetDir()
    {
        $this->sheetDir = $this->dataDir . DIRECTORY_SEPARATOR . 'worksheets';
        $this->createWorkingDir($this->sheetDir);
    }
    
    protected function createBaseDir()
    {
        $this->baseDir = $this->tempDir . DIRECTORY_SEPARATOR . uniqid('ooxml');
        $this->createWorkingDir($this->baseDir);
    }
    
    protected function createContentTypesFile()
    {
        $contentTypesFile = $this->baseDir . DIRECTORY_SEPARATOR . '[Content_Types].xml';
        $this->createWorkingFile($contentTypesFile, '<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>');
    }
    
    protected function createRelsFile()
    {
        $relsDir = $this->baseDir . DIRECTORY_SEPARATOR . '_rels';
        $this->createWorkingDir($relsDir);
        
        $relsFile = $relsDir . DIRECTORY_SEPARATOR . '.rels';
        $this->createWorkingFile($relsFile, '<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>');
    }
    
    protected function createDocPropFiles()
    {
        $docPropDir = $this->baseDir . DIRECTORY_SEPARATOR . 'docProps';
        $this->createWorkingDir($docPropDir);
        
        $appFile = $docPropDir . DIRECTORY_SEPARATOR . 'app.xml';
        $this->createWorkingFile($appFile, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <TotalTime>0</TotalTime>
    <Application>' . self::APP_NAME . '</Application>
</Properties>');
        
        $coreFile = $docPropDir . DIRECTORY_SEPARATOR . 'core.xml';
        $this->createWorkingFile($coreFile, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dcterms:created xsi:type="dcterms:W3CDTF">' . date_create()->format('c') . '</dcterms:created>
    <cp:revision>0</cp:revision>
</cp:coreProperties>');
    }

    protected function zipWorkingFiles()
    {
        
    }
    
    protected function copyZipToStream()
    {
        
    }
    
    protected function cleanWorkingFiles()
    {
        $reversedFiles = array_reverse($this->workingFiles);
        foreach($reversedFiles as $filename) {
            echo $filename . "\n";
            if(is_dir($filename)) {
                if(! rmdir($filename)) {
                    throw new IOException('failed to clean up dir: ' . $filename);
                }
            } else {
                if(! unlink($filename)) {
                    throw new IOException('failed to clean up file: ' . $filename);
                }
            }
        }
    }
    
    protected function createWorkingDir($dir)
    {
        $this->workingFiles[] = $dir;
        if(! mkdir($dir)) {
            throw new IOException('failed to create dir: ' . $dir);
        }
    }
    
    protected function createWorkingFile($filename, $data)
    {
        $this->workingFiles[] = $filename;
        if(! file_put_contents($filename, $data)) {
            throw new IOException('failed to create file: ' . $filename);
        }
    }
    
    protected function createWorkingStream($filename)
    {
        $this->workingFiles[] = $filename;
        return fopen($filename, 'w');
    }
    
    protected function escape($string)
    {
        return htmlspecialchars($string, ENT_QUOTES, 'utf-8');
    }
}