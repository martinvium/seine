<?php
namespace SpreadSheetWriter;

final class ZipArchiveCompressor implements Compressor
{
    public function compressDir($source, $destination)
    {
        $zip = new \ZipArchive();
        $zip->open($destination, \ZIPARCHIVE::CREATE);
        $this->addDir($zip, $source);
        $zip->close();
    }
    
    public function addDir(\ZipArchive $zip, $filename, $localname = null) 
    { 
        $prefix = '';
        if($localname) {
            $prefix = $localname . '/';
        }
        
        $zip->addEmptyDir($localname); 
        $iter = new \RecursiveDirectoryIterator($filename, \FilesystemIterator::SKIP_DOTS); 

        foreach ($iter as $fileinfo) { 
            if (! $fileinfo->isFile() && !$fileinfo->isDir()) { 
                continue; 
            } 

            if($fileinfo->isFile()) {
                $zip->addFile($fileinfo->getPathname(), $prefix . $fileinfo->getFilename());
            } else {
                $this->addDir($zip, $fileinfo->getPathname(), $prefix . $fileinfo->getFilename());
            }
        } 
    } 
}