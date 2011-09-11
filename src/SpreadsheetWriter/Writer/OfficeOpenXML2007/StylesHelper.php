<?php
namespace SpreadSheetWriter\Writer\OfficeOpenXML2007;

use SpreadSheetWriter\Writer\OfficeOpenXML2007StreamWriter as MyWriter;

final class StylesHelper
{
    public function render(array $styles)
    {
        $data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' . MyWriter::EOL;
        $data .= $this->buildStyleFonts($styles);
        $data .= $this->buildFills();
        $data .= $this->buildBorders();
//        $data .= $this->buildCellStyles();
        $data .= $this->buildCellXfs($styles);
        $data .= '</styleSheet>';
        return $data;
    }

    private function buildStyleFonts(array $styles)
    {
        $data = '    <fonts count="' . count($styles) . '">' . MyWriter::EOL;
        foreach($styles as $style) {
            $data .= '        <font>' . MyWriter::EOL;
            if($style->getFontBold()) {
                $data .= '            <b/>' . MyWriter::EOL;
            }
            $data .= '            <sz val="' . ($style->getFontSize() ? $style->getFontSize() : MyWriter::FONT_SIZE_DEFAULT) . '"/>' . MyWriter::EOL;
            $data .= '            <name val="' . ($style->getFontFamily() ? $style->getFontFamily() : MyWriter::FONT_FAMILY_DEFAULT) . '"/>' . MyWriter::EOL;
            $data .= '            <family val="2"/>' . MyWriter::EOL; // no clue why this needs to be there
            $data .= '        </font>' . MyWriter::EOL;
        }
        $data .= '    </fonts>' . MyWriter::EOL;
        return $data;
    }

    private function buildFills()
    {
        return '    <fills count="1">
        <fill>
            <patternFill patternType="none"/>
        </fill>
    </fills>';
    }

    private function buildBorders()
    {
        return '    <borders count="1">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
    </borders>';
    }

    private function buildCellStyles()
    {
        return '<cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>';
    }

    private function buildCellXfs(array $styles)
    {
        $i = 0;
        $data = '    <cellXfs count="' . count($styles) . '">' . MyWriter::EOL;
        foreach($styles as $style) {
            $data .= '        <xf numFmtId="0" fontId="' . $i . '" fillId="0" borderId="0" xfId="0" applyFont="' . ($i > 0 ? 1 : 0) . '"/>' . MyWriter::EOL;
            $i++;
        }
        $data .= '    </cellXfs>' . MyWriter::EOL;
        return $data;
    }
}