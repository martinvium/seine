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

/**
 * Shared configuration class for SpreadSheetWriter. Generic option consts are found on this class.
 * Options specific to each writer, can be found on each writers base implementation class.
 *
 * @author Martin Vium <code@codefuss.com>
 * @example $config->setOption(Configuration::OPT_TEMP_DIR, '/tmp');
 * @example $config->setOption(CSVWriter::OPT_TEXT_DELIMITER, ';');
 */
class Configuration
{
    const OPT_WRITER = 'Configuration::OPT_WRITER';
    const OPT_TEMP_DIR = 'Configuration::OPT_TEMP_DIR';

    private $options = array();

    public function setOption($name, $value)
    {
        $this->options[$name] = $value;
    }

    public function setOptions(array $options)
    {
        $this->options = array_merge($this->options, $options);
    }

    public function getOption($name, $default = null)
    {
        if(! isset($this->options[$name])) {
            return $default;
        }

        return $this->options[$name];
    }
}