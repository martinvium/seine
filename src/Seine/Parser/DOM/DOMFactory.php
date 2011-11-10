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
namespace Seine\Parser\DOM;

use Seine\Factory;
use Seine\Writer\WriterFactoryImpl;
use Seine\Writer;
use Seine\Configuration;
use Seine\Book;

/**
 * Base of operation when creating spreadsheets.
 * @author Martin Vium
 */
final class DOMFactory implements Factory 
{
    /**
     * @var Configuration
     */
    private $config;

    /**
     * Create a factory object specifying the default configuration.
     *
     * @param Configuration $config
     * @return DOMFactory
     */
    public static function FromConfig(Configuration $config)
    {
        $factory = new self();
        $factory->setConfig($config);
        return $factory;
    }

    public function __construct()
    {
        $this->config = new Configuration;
    }

    /**
     * Set the default configuration for getConfiguredWriter() and getConfiguredBook().
     *
     * @param Configuration $config
     */
    public function setConfig(Configuration $config)
    {
        $this->config = $config;
    }
    
    /**
     * @return WriterFactoryImpl 
     */
    public function getWriterFactory()
    {
        return new WriterFactoryImpl($this);
    }
    
    /**
     * Get a "clean" book object, you must set a writer before using it.
     *
     * @see Book::setWriter()
     * @return DOMBook 
     */
    public function getBook()
    {
        return new DOMBook($this);
    }
    
    /**
     * Get a new row object.
     *
     * @param array $cells
     * @return DOMArrayRow
     */
    public function getRow(array $cells)
    {
        return new DOMArrayRow($this, $cells);
    }
    
    /**
     * @internal
     * @return DOMSheet 
     */
    public function getSheet()
    {
        return new DOMSheet($this);
    }

    /**
     * @internal
     * @see Book::newStyle()
     * @param string $id
     * @return DOMStyle
     */
    public function getStyle($id)
    {
        return new DOMStyle($this, $id);
    }
    
    /**
     * Get a Book already configured with a writer and options from the default configuration
     * or the $config param if passed in.
     *
     * @param stream $fp
     * @param Configuration $config
     * @return Book
     */
    public function getConfiguredBook($fp, Configuration $config = null)
    {
        $book = $this->getBook();
        $book->setWriter($this->getConfiguredWriter($fp, $config));
        return $book;
    }

    /**
     * Get a writer configured using the default configuration or the one passed in.
     *
     * @param stream $fp
     * @param Configuration $config
     * @return Writer
     */
    public function getConfiguredWriter($fp, Configuration $config = null)
    {
        if(! $config) {
            $config = $this->config;
        }

        $writerName = $config->getOption(Configuration::OPT_WRITER);
        if(! $writerName) {
            throw new \Exception('Writer must be defined in config for getConfiguredWriter()');
        }

        $writer = $this->getWriterByName($fp, $writerName);
        $writer->setConfig($config);
        return $writer;
    }

    /**
     * Get a writer by it's name. The name is looked up in the writer factory prefixed with "get".
     *
     * @example $writer = $factory->getWriterByName($fp, 'OfficeOpenXML2003StreamWriter');
     * @param stream $fp
     * @param string $writerName
     * @return Writer
     */
    public function getWriterByName($fp, $writerName)
    {
        $method = 'get' . $writerName;
        if(method_exists($this->getWriterFactory(), $method)) {
            return $this->getWriterFactory()->$method($fp);
        } else {
            throw new \InvalidArgumentException('writer not found: ' . $writerName);
        }
    }
}