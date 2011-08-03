<?php
namespace SpreadSheetWriter\Parser\DOM;

use SpreadSheetWriter\Factory;
use SpreadSheetWriter\Writer;
use SpreadSheetWriter\Element;

abstract class DOMElement implements Element
{
    /**
     * @var Factory
     */
    protected $factory;
    
    public function __construct(Factory $factory)
    {
        $this->factory = $factory;
    }
}