<?php
namespace SpreadSheetWriter\Writer\OfficeOpenXML2007;

final class Utils
{
    static public function escape($string)
    {
        return htmlspecialchars($string, ENT_QUOTES, 'utf-8');
    }
}