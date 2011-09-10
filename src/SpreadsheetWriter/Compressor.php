<?php
namespace SpreadSheetWriter;

/**
 * @author Martin Vium
 */
interface Compressor
{
    public function compressDir($source, $destination);
}