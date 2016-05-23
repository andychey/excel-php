<?php

namespace Andychey\Excel;

use PHPExcel;


class Reader
{
    /**
     * 文件类型
     *
     * @var string
     */
    protected static $type = 'xlsx';

    /**
     * 文件名
     *
     * @var
     */
    protected static $filename;

    /**
     * 要忽略的行
     *
     * @var
     */
    protected static $skip_rows = array();

    /**
     * PHPExcelReader对象
     *
     * @var
     */
    protected static $phpExcelReader;

    /**
     * PHPExcel对象
     *
     * @var PHPExcel
     */
    protected static $phpExcel;

    /**
     * @param string $filename
     * @param null $skip_rows
     * @param null $type
     *
     * @return array
     */
    public static function loadToArray($filename, $skip_rows = null, $type = null)
    {
        if (! file_exists($filename)) {
            throw new \InvalidArgumentException("File 「{$filename}」 does't exist");
        }
        self::$filename = $filename;
        if (! is_null($type)) {
            self::$type = $type;
        }
        if (! is_null($skip_rows)) {
            if (! is_array($skip_rows)) {
                $skip_rows = array($skip_rows);
            }
            self::$skip_rows = $skip_rows;
        }
        self::loadToPHPExcel();
        return array_diff_key(self::$phpExcel->getActiveSheet()->toArray(), array_flip(self::$skip_rows));
    }

    /**
     * 将文件加载至PHPExcel对象中
     */
    protected static function loadToPHPExcel()
    {
        self::$phpExcelReader = self::createReader();
        if (self::$skip_rows) {
            self::$phpExcelReader->setReadFilter(new ReadFilter(self::$skip_rows));
        }
        self::$phpExcel = self::$phpExcelReader->load(self::$filename);
    }

    /**
     * 创建阅读器
     *
     * @return \PHPExcel_Reader_IReader
     */
    protected static function createReader()
    {
        new PHPExcel();
        switch (self::$type) {
            case 'xls':
                $reader = \PHPExcel_IOFactory::createReader('Excel5');
                break;
            case 'xlsx':
            default:
                $reader = \PHPExcel_IOFactory::createReader('Excel2007');
                break;
        }
        return $reader;
    }
}