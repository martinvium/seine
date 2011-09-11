<?php
/**
 * Copyright (C) 2011 by Martin Vium
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */
namespace SpreadSheetWriter\Writer\OfficeOpenXML2007;

use SpreadSheetWriter\Writer;
use SpreadSheetWriter\IOException;
use SpreadSheetWriter\Compressor;

abstract class WriterBase implements Writer
{
    const APP_NAME = 'SpreadSheetWriter';
    const EOL = "\r\n";
    
    /**
     * @var Compressor
     */
    private $compressor;
    
    private $workingFiles = array();
    
    protected $stream;
    protected $tempDir;
    protected $baseDir;
    protected $dataDir;
    protected $sheetDir;
    protected $dataRelsDir;
    
    public function __construct($stream, Compressor $compressor)
    {
        if(! is_resource($stream)) {
            throw new \InvalidArgumentException('fp is not a valid stream resource');
        }
        
        $this->stream = $stream;
        $this->compressor = $compressor;
        $this->tempDir = sys_get_temp_dir();
    }
    
    public function setTempDir($dir)
    {
        $this->tempDir = $dir;
    }
    
    protected function createBaseStructure()
    {
        $this->createBaseDir();
        $this->createRelationsFile();
        $this->createDocPropFiles();
        $this->createDataDir();
        $this->createSheetDir();
        $this->createDataRelsDir();
    }
    
    private function createDataRelsDir()
    {
        $this->dataRelsDir = $this->dataDir . DIRECTORY_SEPARATOR . '_rels';
        $this->createWorkingDir($this->dataRelsDir);
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
    
    private function createRelationsFile()
    {
        $relsDir = $this->baseDir . DIRECTORY_SEPARATOR . '_rels';
        $this->createWorkingDir($relsDir);
        
        $relationsFile = $relsDir . DIRECTORY_SEPARATOR . '.rels';
        $this->createWorkingFile($relationsFile, '<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rIdWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rIdApp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
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
        $this->compressor->compressDir($this->baseDir, $this->baseDir . '.zip');
    }

    protected function copyZipToStream()
    {
        $source = fopen($this->baseDir . '.zip', 'r');
        stream_copy_to_stream($source, $this->stream);
        fclose($source);
        if(! unlink($this->baseDir . '.zip')) {
            throw new IOException('failed to clean up file: ' . $source);
        }
    }
    
    protected function cleanWorkingFiles()
    {
        $reversedFiles = array_reverse($this->workingFiles);
        foreach($reversedFiles as $filename) {
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
            throw new IOException('failed to create file and put data: ' . $filename);
        }
    }

    protected function createEmptyWorkingFile($filename)
    {
        $this->workingFiles[] = $filename;
        if(! touch($filename)) {
            throw new IOException('failed to create empty file: ' . $filename);
        }
    }
    
    protected function createWorkingStream($filename)
    {
        $this->workingFiles[] = $filename;
        return fopen($filename, 'w');
    }
    
    protected function escape($string)
    {
        return Utils::escape($string);
    }
}