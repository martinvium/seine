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
namespace SpreadSheetWriter;

interface Sheet
{
    /**
     * @internal assigned by Book
     * @access private
     */
    public function setBook(Book $book);

    /**
     * Add a row to the sheet
     *
     * @param Row $row
     * @return Sheet
     */
    public function addRow(Row $row);

    /**
     * Get the visible name of the sheet
     *
     * @return string
     */
    public function getName();
    
    /**
     * Get the internal unique id of  the sheet
     *
     * @access private
     * @return string
     */
    public function getId();
    
    /**
     * Used by Book to assign a unique id to the sheet.
     *
     * @access private
     * @return Sheet
     */
    public function setId($id);
    
    public function setName($name);
}