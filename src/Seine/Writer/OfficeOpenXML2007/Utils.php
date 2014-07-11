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

final class Utils
{
    /** @var string[] Control characters to be escaped */
    private static $controlCharactersSet;

    /**
     * Escapes the given string to make it compatible with Excel
     *
     * @param string $string The string to be escaped
     * @return string
     */
    public static function escape($string)
    {
        $escapedString = self::escapeControlCharacters($string);
        $escapedString = htmlspecialchars($escapedString, ENT_QUOTES, 'utf-8');

        return $escapedString;
    }

    /**
     * Returns the control characters array. Builds it once and cache it.
     *
     * @return string[]
     */
    private static function getControlCharactersSet()
    {
        if (!self::$controlCharactersSet) {
            // Build the array if not already built
            for ($i = 0; $i <= 31; ++$i) {
                if ($i != 9 && $i != 10 && $i != 13) {
                    $escapedValue = '_x' . sprintf('%04s' , strtoupper(dechex($i))) . '_';
                    $rawValue = chr($i);
                    self::$controlCharactersSet[$escapedValue] = $rawValue;
                }
            }
        }

        return self::$controlCharactersSet;
    }

    /**
     * Converts PHP control characters from the given string to OpenXML escaped control characters
     *
     * Control characters are stored directly in the shared-strings table.
     * Characters that cannot be represented in XML are encoded using the following escape sequence:
     * _xHHHH_ where H represents a hexadecimal character in the character's value...
     * So you could end up with something like _x0008_ in a string (either in a cell value (<v>)
     * element or in the shared string <t> element.
     *
     * @param string $string String to escape
     * @return string
     */
    private static function escapeControlCharacters($string)
    {
        $controlCharactersSet = self::getControlCharactersSet();
        return str_replace(array_values($controlCharactersSet), array_keys($controlCharactersSet), $string);
    }
}