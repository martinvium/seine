<?php
$src_path = dirname(__DIR__) . '/src';
$tests_path = __DIR__;
set_include_path(get_include_path() . PATH_SEPARATOR . $src_path . PATH_SEPARATOR . $tests_path);

spl_autoload_register(function($class) {
    return spl_autoload(str_replace('\\', '/', $class));
});

require_once 'PHPUnit/Framework/TestCase.php';