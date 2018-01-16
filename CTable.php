<?php
/**
 * Created by PhpStorm.
 * User: baishancarrot
 * Date: 2017/12/12
 * Time: 下午5:33
 */

namespace stone\lib;   //依据具体项目修改

class CTable
{
    CONST NORMAL_ARRAY_TYPE = 0;
    CONST MERGE_ROW_TYPE = 1;

    private $objPHPExcel;
    private $arHead;
    private $arData;
    private $merge;
    private $mergeCount;
    private $num = 0;

    /* excel通用导出类
     | ——————————————————
     | 使用场景
     | ——————————————————
     | 特殊的合并单元
     | 多个sheet导出
     | 一个单元格输出多行数据
     | 单元格样式有要求（行高、颜色、对齐方式等）
     | ——————————————————
     | 如果需要导出csv格式数据
     | 或者简单的导出excel数据
     | 请直接使用static方法
     | ——————————————————
     */
    public function __construct()
    {
        //从新设置缓存路径
        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_in_memory_gzip;
        $cacheSettings = [];
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);
        //实例化类
        $this->objPHPExcel = new \PHPExcel();
    }

    /*
     * 增加excel页数
     | ——————————————
     | 数据格式
     | ——————————————
     |
     |
     | ——————————————
     */
    public function addSheet($arRet, $type, $sheetName = null)
    {
        //创建新的sheet
        $this->objPHPExcel->createSheet($this->num);
        //切换工作sheet
        $this->objPHPExcel->setActiveSheetIndex($this->num);
        //设置当前工作sheet名称
        $sheetName = isset($sheetName) ? $sheetName : 'Page ' . $this->num;
        $this->objPHPExcel->getActiveSheet()->setTitle($sheetName);
        //初始化sheet
        $this->sheetInit();

        //根据类型处理单元格
        switch ($type) {
            case self::NORMAL_ARRAY_TYPE:
                $this->outXlsXlsx($arRet);
                break;
            case self::MERGE_ROW_TYPE:
                $this->rowMerge($arRet);
                break;
        }

        $this->num++;
    }

    //初始化sheet
    private function sheetInit()
    {
        $this->arHead = [];
        $this->arData = [];
        $this->merge = [];
        $this->mergeCount = [];
    }

    //基础xls输出方法 type = 0
    private function outXlsXlsx($arRet)
    {
        //写入表头
        $this->setColumns($arRet['columns']);
        //写入主要数据
        $this->arData = $arRet['datas'];
        $this->setDatas();
    }

    //行合并通用方法 type = 1
    private function rowMerge($arRet)
    {
        //写入表头
        $this->setColumns($arRet['columns']);
        //处理合并
        $this->setMerges($arRet['datas']);
        //解合并
        $this->mergeCell();
        //写入主要数据
        $this->setDatas();
    }

    //输出excel
    public function outExcel($filename)
    {
        header("Cache-Control:no-cache,must-revalidate");
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $filename . '.xls"');
        header('Cache-Control: max-age=0');
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0
        $objWriter = \PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
    }

    //输入表头
    private function setColumns($columns)
    {
        $i = 0;
        $return = [];
        foreach ($columns as $column) {
            //单个key
            if (empty($column['keys'])) {
                $title = $column['title'];
                $this->writeIntoExcel($title, 1, $i);
                $i++;
                $return[] = $column;
                //设置样式
                if (!empty($column['style'])) {
                    $this->setStyle($column['style'], $i, 1);
                }
            } else { //keys组
                $title = $column['title'];
                $this->writeIntoExcel($title, 1, $i);
                $count = count($column['keys']);
                if ($count > 1) {
                    $this->mergeCells(1, $i, 1, $i + $count - 1);
                }
                foreach ($column['keys'] as $key) {
                    $key['title'] = $title;
                    $return[] = $key;
                    if (!empty($key['style'])) {
                        $this->setStyle($key['style'], $i, 1);
                    }
                    $i++;
                }
            }
        }

        $this->arHead = $return;
    }

    //输入表体
    private function setDatas()
    {
        $datas = $this->arData;
        $headers = $this->arHead;
        //输入表体
        $rows = 2;
        foreach ($datas as $data) {
            $col = 0;
            foreach ($headers as $header) {
                $key = $header['key'];
                $cellData = null;
                if (isset($data[$key])) {
                    $cellData = $data[$key];
                }
                $this->writeIntoExcel($cellData, $rows, $col);
                //如果设置了扩展函数
                if (isset($header['vendor'])) {
                    $this->vendor($header['vendor'], $col, $rows, $data);
                }
                $col++;
            }
            $rows++;
        }
    }

    //处理扩展函数
    public function vendor($vendor, $col, $row, $data)
    {
        $result = $vendor($col, $row, $data);

        //合并回调
        if (!empty($result['merge'])) {
            list($row1, $col1, $row2, $col2) = $result['merge'];
            $this->mergeCells($row1, $col1, $row2, $col2);
        }
        //如果设置原生样式，则直接使用原生样式
        if (!empty($result['primitive'])) {
            $this->primitive($col, $row, $result['primitive']);

            return;
        }
        //特殊样式回调
        if (!empty($result['style'])) {
            $this->setStyle($result['style'], $col, $row);

            return;
        }
    }

    //设置表格样式
    private function setStyle($style, $col, $row = null)
    {
        //列宽样式
        if (isset($style['width'])) {
            if ($style['width'] == 'auto') {
                $this->setColumnAutoWidth($col);
            } else {
                $this->setColumnWidth($col, $style['width']);
            }
        }
        //字体颜色
        if (isset($style['color'])) {
            $this->setCellFontColor($row, $col, $style['color']);
        }
        //对齐方式
        if (isset($style['align'])) {
            $this->setCellHorizontal($row, $col, $style['align']);
        }
        //设置行高
        if (isset($style['height'])) {
            $this->setHeight($row, $style['height']);
        }
    }

    //处理合并单元格
    private function setMerges($datas)
    {
        $this->arData = $datas;
        $headers = $this->arHead;
        $datas = $this->arData;
        //按照排序顺序获取合并项
        $merges = [];
        foreach ($headers as $header) {
            if (isset($header['merge']) && $header['merge'] && isset($header['sort'])) {
                $merges[$header['sort']] = $header['key'];
            }
        }
        ksort($merges);
        foreach ($headers as $header) {
            if (isset($header['merge']) && $header['merge'] && !isset($header['sort'])) {
                $merges[] = $header['key'];
            }
        }

        $this->merge = $merges;

        $datas = $this->togetherByKey($datas, 0);
        $this->arData = $datas;
    }

    //递归合并
    private function togetherByKey($array, $level)
    {
        //合到最后一层后直接返回数组
        if (empty($this->merge[$level])) {
            return $array;
        }
        $merges = [];
        $key = $this->merge[$level];
        foreach ($array as $data) {
            if (empty($data[$key])) {
                continue;
            }
            $title = $data[$key];
            $merges[$title][] = $data;
        }
        $return = [];
        foreach ($merges as $title => $arr) {
            $return[$title] = $this->togetherByKey($arr, $level + 1);
        }

        return $return;
    }

    //解array合并单元格
    private function mergeCell()
    {
        $datas = $this->arData;
        $this->arData = [];
        $this->disivionArray($datas, 0, 2);
        $mergeCount = $this->mergeCount;

        //获取所有需要合并项
        foreach ($this->arHead as $head) {
            $key = $head['key'];
            if (empty($head['merge']) || !$head['merge']) {
                continue;
            }
            if (empty($head['affect'])) {
                continue;
            }
            foreach ($head['affect'] as $affect) {
                $mergeCount[$affect] = $mergeCount[$key];
            }
        }

        $col = -1;
        foreach ($this->arHead as $head) {
            $col++;
            $key = $head['key'];
            if (empty($mergeCount[$key])) {
                continue;
            }
            foreach ($mergeCount[$key] as $merge) {
                $start = $merge['start'];
                $end = $merge['end'];
                if ($start == $end) {
                    continue;
                }
                $this->mergeCells($start, $col, $end, $col);
            }
        }
    }

    //递归解array
    private function disivionArray($array, $level, $row)
    {
        if (empty($this->merge[$level])) {
            $this->arData = empty($this->arData) ? $array : array_merge($this->arData, $array);

            return count($array);
        }

        $key = $this->merge[$level];
        $rowCount = 0;
        foreach ($array as $arr) {
            $start = $row;
            $count = $this->disivionArray($arr, $level + 1, $row);
            if ($count == 0) {
                continue;
            }
            $row += $count;
            $end = $row - 1;
            $rowCount += $count;
            //获取应合并的行列
            $this->mergeCount[$key][] = [
                'start' => $start,
                'end'   => $end,
            ];
        }

        return $rowCount;
    }

    //写入excel对应行列
    private function writeIntoExcel($value, $row, $col)
    {
        if (!isset($value)) {
            return;
        }
        $index = $this->getColumnName($col) . $row;
        //如果数据是array则多行输出
        if (is_array($value)) {
            $str = null;

            foreach ($value as $string) {
                $string = isset($string) ? $string : ' ';
                $str = isset($str) ? $str . chr(10) . $string : $string;
            }
            $value = isset($str) ? $str : ' ';

            //设置自动换行
            $this->objPHPExcel->getActiveSheet()->getStyle($index)->getAlignment()->setWrapText(true);
        }
        //写入数据
        $this->objPHPExcel->getActiveSheet()->setCellValue($index, $value);
    }

    //合并单元格
    private function mergeCells($row1, $col1, $row2, $col2)
    {
        $this->objPHPExcel->getActiveSheet()
            ->mergeCells(
                $this->getColumnName($col1) . $row1 . ':' .
                $this->getColumnName($col2) . $row2
            );
    }

    //根据列数获取列名
    private function getColumnName($col)
    {
        $name = '';
        while ($col >= 0) {
            $num = $col % 26;
            $name = chr($num + ord("A")) . $name;
            $col = intval(($col - $num) / 26) - 1;
        }

        return $name;
    }

    /*样式设置函数*/

    //设置列宽自动适应
    private function setColumnAutoWidth($col)
    {
        $this->objPHPExcel->getActiveSheet()
            ->getColumnDimension($this->getColumnName($col))->setAutoSize(true);
    }

    //设置列宽
    private function setColumnWidth($col, $width)
    {
        $this->objPHPExcel->getActiveSheet()
            ->getColumnDimension($this->getColumnName($col))->setWidth($width);
    }

    //单元格字体颜色
    private function setCellFontColor($row, $col, $color)
    {
        $this->objPHPExcel->getActiveSheet()->getStyle($this->getColumnName($col) . $row)
            ->getFont()->getColor()->setRGB($color);
    }

    //设置对齐方式
    private function setCellHorizontal($row, $col, $align)
    {
        $this->objPHPExcel->getActiveSheet()->getStyle($this->getColumnName($col) . $row)
            ->getAlignment()->setHorizontal($align);
    }

    //直接使用原生样式
    private function primitive($row, $col, $primitive)
    {
        $this->objPHPExcel->getActiveSheet()->getStyle($this->getColumnName($col) . $row)
            ->applyFromArray($primitive);
    }

    //设置行高
    private function setHeight($row, $height)
    {
        $this->objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight($height);
    }

    /* * * * * * * * *
     * 常用static方法 *
     * * * * * * * * */

    /*
     * 导出格式csv
     * */
    public static function outCsv($fileName, $headers = [], $datas = [], $default = null)
    {
        //当单元格没有相应数据的默认值，默认输出空格
        $default = empty($default) ? ' ' : $default;

        //计算表头
        $headerData = [];
        foreach ($headers as $header) {
            $headerData[] = $header['title'];
        }
        //计算表体
        $bodyData = [];
        foreach ($datas as $data) {
            $rowData = [];
            foreach ($headers as $header) {
                $key = $header['key'];
                $string = $default;
                if (isset($data[$key])) {
                    $string = $data[$key];
                }
                array_push($rowData, $string);
            }

            $bodyData[] = $rowData;
        }

        self::outSimpleCsv($fileName, $headerData, $bodyData);
    }

    /*
     * 导出简单的csv数据
     * */
    public static function outSimpleCsv($fileName, $header = [], $datas = [])
    {
        header('Content-Type: application/octet-stream');
        header('Content-Disposition: attachment; filename="' . $fileName . '.csv"');
        if (!empty($header)) {
            echo iconv('utf-8', 'gbk//TRANSLIT', '"' . implode('","', $header) . '"' . "\n");
        }
        foreach ($datas as $key => $value) {
            echo iconv('utf-8', 'gbk//TRANSLIT', '"' . implode('","', $value) . "\"\n");
        }
    }

}
