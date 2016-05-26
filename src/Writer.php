<?php

namespace Andychey\Excel;

use PHPExcel;
use PHPExcel_IOFactory;

use InvalidArgumentException;


class Writer
{
    /**
     * 表头
     *
     * @var array
     */
    protected $head = array();

    /**
     * 表头样式
     *
     * @var array
     */
    protected $headStyle = array(
        'font' => array('bold' => true),
        'alignment' => array('horizontal' => 'center')
    );

    /**
     * 数据
     *
     * @var array
     */
    protected $data = array();

    /**
     * 文件类型
     *
     * @var string
     */
    protected $type = 'xlsx';

    /**
     * PHPExcel对象
     *
     * @var PHPExcel
     */
    protected $phpExcel;

    /**
     * 列序号
     */
    protected static $column_serial = array(
        'A','B','C','D','E','F','G', 'H','I','J',
        'K','L','M','N', 'O','P','Q','R','S','T',
        'U', 'V','W','X','Y','Z','AA','AB','AC','AD',
        'AE', 'AF','AG','AH','AI','AJ','AK','AL','AM','AN',
        'AO', 'AP','AQ','AR','AS','AT','AU','AV','AW','AX',
    );

    /**
     * 最大数字宽度
     */
    const MAX_NUMBER_WIDTH = 11;

    /**
     * Writer constructor
     *
     * @param array      $head
     * @param array|null $data
     * @param null       $type
     */
    public function __construct(array $head, array $data = null, $type = null)
    {
        $this->head = $head;
        if (! is_null($data)) {
            $this->data = $data;
        }
        if (! is_null($type)) {
            $this->type = $type;
        }
        $this->phpExcel = new PHPExcel();
        $this->init();
    }

    /**
     * 直接输出到浏览器（下载）
     *
     * @param $filename
     */
    public function download($filename)
    {
        $this->process();

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0');

        $write = $this->createWriter();
        $write->save('php://output');
    }

    /**
     * 保存为文件
     *
     * @param $filename
     */
    public function save($filename)
    {
        $this->process();
        $write = $this->createWriter();
        $write->save($filename);
    }

    /**
     * 设置表头样式
     *
     * @param array $style
     */
    public function setHeadStyle(array $style)
    {
        $this->headStyle = array_merge($this->headStyle, $style);
    }

    /**
     * 设置文件类型
     *
     * @param $type
     */
    public function setType($type)
    {
        $this->type = $type;
    }

    /**
     * 添加一行数据
     * 
     * @param array $row
     *
     * @throws Exception
     */
    public function addRow(array $row)
    {
        if (count($row) != count($this->head)) {
            throw new InvalidArgumentException("The row does't match this head");
        }
        $this->data[] = $row;
    }

    /**
     * 处理数据
     */
    protected function process()
    {
        $this->processHead();
        $this->processData();
    }

    /**
     * 处理表头
     */
    protected function processHead()
    {
        $last_column = self::$column_serial[count($this->head) - 1];
        $actSheet = $this->phpExcel->getActiveSheet();
        $actSheet->getStyle("A1:{$last_column}1")->applyFromArray($this->headStyle);
        foreach ($this->head as $serial => $head) {
            $column = self::$column_serial[$serial];
            $actSheet->getColumnDimension($column)->setAutoSize();
            $cell = "{$column}1";
            if (is_array($head)) {
                if (isset($head['style'])) {
                    $actSheet->getStyle($cell)->applyFromArray($head['style']);
                }
                if (isset($head['width'])) {
                    $actSheet->getColumnDimension($column)->setWidth($head['width']);
                }
                $value = isset($head['value']) ? $head['value'] : '';
            } else {
                $value = $head;
            }
            if (is_scalar($value) && null !== $value) {
                $actSheet->setCellValue($cell, $value);
            }
        }
    }

    /**
     * 处理数据
     */
    protected function processData()
    {
        $actSheet = $this->phpExcel->getActiveSheet();
        $row_num = 2;
        foreach ($this->data as $row) {
            foreach ($row as $serial => $data) {
                $column = self::$column_serial[$serial];
                $cell = "{$column}{$row_num}";
                if (isset($this->head[$serial]['dataStyle'])) {
                    $actSheet->getStyle($cell)->applyFromArray($this->head[$serial]['dataStyle']);
                }
                if (is_array($data)) {
                    if (isset($data['style'])) {
                        $actSheet->getStyle($cell)->applyFromArray($data['style']);
                    }
                    $value = isset($data['value']) ? $data['value'] : null;
                } else {
                    $value = $data;
                }
                if (is_scalar($value) && null !== $value) {
                    if (is_numeric($value) && strlen($value) > self::MAX_NUMBER_WIDTH) {
                        $value = '`' . $value;
                    }
                    $actSheet->setCellValue($cell, $value);
                }
            }
            $row_num ++;
        }
    }

    /**
     * 初始化
     */
    protected function init()
    {
        return $this->phpExcel
            ->getActiveSheet()
            ->getDefaultRowDimension()
            ->setRowHeight(25);
    }

    /**
     * 创建抄写员
     *
     * @return PHPExcel_Writer_IWriter
     */
    protected function createWriter()
    {
        switch ($this->type) {
            case 'csv':
                $writer = PHPExcel_IOFactory::createWriter($this->phpExcel, 'CSV');
                break;
            case 'xls':
                $writer = PHPExcel_IOFactory::createWriter($this->phpExcel, 'Excel5');
                break;
            case 'xlsx':
            default:
                $writer = PHPExcel_IOFactory::createWriter($this->phpExcel, 'Excel2007');
        }
        return $writer;
    }
}