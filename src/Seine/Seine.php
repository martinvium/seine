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
namespace Seine;

use Seine\Parser\DOM\DOMFactory;
use Seine\Configuration;

class Seine
{
  private $config;
  private $factory;

  private $writerMap = array(
    'csv' => 'CSVStreamWriter'
  );

  public function __construct($options = array()) 
  {
    $this->config = new Configuration();
    $this->config->setOption(Configuration::OPT_WRITER, $this->getWriterOption($options));
    $this->factory = new DOMFactory();
  }

  /**
   * @return Book
   */
  public function newDocument($filename, $mode = 'w') 
  {
    $fp = fopen($filename, $mode);
    $config = clone $this->config;
    $config->setOption(Configuration::OPT_AUTO_CLOSE_STREAM, true);
    return $this->factory->getConfiguredBook($fp, $config);
  }

  /**
   * @return Book
   */
  public function newDocumentFromStream($fp) {
    return $this->factory->getConfiguredBook($fp, $this->config);
  }

  public function setOption($name, $value)
  {
      $this->config->setOption($name, $value);
  }

  private function getWriterOption($options) 
  {
    $value = $options['writer'];
    return $this->writerMap[$value];
  }
}