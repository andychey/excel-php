<?php

namespace Andychey\Excel;

use PHPUnit_Framework_TestCase;


class ReaderTest extends PHPUnit_Framework_TestCase
{
    public function testRead()
    {
        $filename_src = __DIR__ . '/data/src.xlsx';

        $data = Reader::loadToArray($filename_src, 0);

        $this->assertNotEmpty($data);
    }
}