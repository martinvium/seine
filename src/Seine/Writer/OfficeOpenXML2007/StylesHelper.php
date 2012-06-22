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
namespace Seine\Writer\OfficeOpenXML2007;

use Seine\Writer\OfficeOpenXML2007StreamWriter as MyWriter;

final class StylesHelper
{
    const FONT_FAMILY_DEFAULT = 'Arial';
    const FONT_SIZE_DEFAULT = 10;

    public function render(array $styles)
    {
        $data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' . MyWriter::EOL;
        $data .= $this->buildNumFmts();
        $data .= $this->buildStyleFonts($styles);
        $data .= $this->buildFills();
        $data .= $this->buildBorders();
        $data .= $this->buildCellStyles();
        $data .= $this->buildCellStyleXfs();
        $data .= $this->buildCellXfsDefault();
//        $data .= $this->buildCellXfs($styles);
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
            $data .= '            <sz val="' . ($style->getFontSize() ? $style->getFontSize() : self::FONT_SIZE_DEFAULT) . '"/>' . MyWriter::EOL;
            $data .= '            <name val="' . ($style->getFontFamily() ? $style->getFontFamily() : self::FONT_FAMILY_DEFAULT) . '"/>' . MyWriter::EOL;
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

    private function buildCellStyleXfs()
    {
        return '<cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="false" applyFont="true" applyBorder="false" applyAlignment="false"/>
    </cellStyleXfs>';
    }

    private function buildCellStyles()
    {
        return '<cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>';
    }

    private function buildNumFmts()
    {
        return '    <numFmts count="5">
        <numFmt numFmtId="164" formatCode="GENERAL"/>
        <numFmt numFmtId="165" formatCode="DD/MM/YY"/>
        <numFmt numFmtId="166" formatCode="0%"/>
        <numFmt numFmtId="167" formatCode="HH:MM:SS"/>
        <numFmt numFmtId="168" formatCode="DD/MM/YY\ HH:MM"/>
    </numFmts>';
    }

    private function buildCellXfsDefault()
    {
        $data = '    <cellXfs count="5">' . MyWriter::EOL;
        for($i = 164; $i <= 168; $i++) {
            $data .= '        <xf numFmtId="' . $i . '" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="false"/>' . MyWriter::EOL;
        }
        $data .= '    </cellXfs>' . MyWriter::EOL;
        return $data;
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