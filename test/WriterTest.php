<?php

namespace Andychey\Excel;


class WriterTest extends \PHPUnit_Framework_TestCase
{
    public function testSave()
    {
        $head = array(
            array('value'=>'商品ID', 'width'=>10),
            array('value'=>'用户Uin', 'width'=>25),
            array('value'=>'收货人', 'width'=>20),
            array('value'=>'手机号', 'width'=>20),
            array('value'=>'收货地址', 'width'=>50),
        );
        $row = array(
            '126', '148461050718576940','十二OK您','15998568584','北京 北京市 延庆县 您人痛快淋漓今婆婆和和气气'
        );

        $writer = new Writer($head);

        $writer->addRow($row);

        $filename_test = __DIR__ . '/data/out.xlsx';

        $writer->save($filename_test);

        $this->assertFileExists($filename_test);
    }
}