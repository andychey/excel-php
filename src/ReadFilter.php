<?php

namespace Andychey\Excel;


class ReadFilter implements \PHPExcel_Reader_IReadFilter
{
    /**
     * 要忽略的行
     *
     * @var array
     */
    protected $skip_rows = array();

    /**
     * ReadFilter constructor
     *
     * @param $skip_rows
     */
    public function __construct($skip_rows)
    {
        $this->skip_rows = $skip_rows;
    }

    /**
     * @param        $column
     * @param        $row
     * @param string $worksheetName
     *
     * @return bool
     */
    public function readCell($column, $row, $worksheetName = '')
    {
        if (in_array($row - 1, $this->skip_rows)) {
			return false;
		}
		return true;
    }
}