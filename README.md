# kaiPHP.github.io
kaiPHP个人博客

/**封装php导出excel方法适用多种样式

/**
 * Excel 导出excel 到本地
 * @param array $data 要导出的数据    $data['fields'] 为导出表头的名称及样式 $data['dataList']为要导出的数据
 * @param string $filename 文件名称
 * @param array $mergeCells 待合并的单元格，如：Array ('A3:A8', 'B3:B8')， 可不传
 * @param array $style 设置表头信息 颜色、背景颜色、表格高度、表格宽度、表格线 可不传
 * @param string $sheetname Excel    sheet名 可不传
 * @param object $objectPHPExcel Excel    Excel对象可不传
 * @param array $sheetIndex 当前操作的sheet 可不传
 * @author yangkai 2018/8/15
 * $style = array(
 *      'heigh'=>30,（表头高度）
 *      'width'=>20,(表头宽度)
 *      'line'=>0（0，有表格线，默认有表格线，1，无表格线）
 *      以上参数均可填可不填
 * );
 * $data['fields'] = array(
 * array( //第一行表头样式，多个配置多个数组
 * array(
 * 'merge' => 9,                   //向后合并9个单元格
 * 'value' => '代理商开通情况',      //单元格内容
 * 'remove_ver_align' => '1',      //垂直居中(默认居中)
 * 'remove_text_align' => '1',     //水平居中(默认居中)
 * 'color' => 'FFFFFF',            //文字颜色
 * 'background' => '00E5EE',       //单元格背景颜色
 * 'width' => '30',                //设置所有列的宽度
 * 'height' => '30',               //设置所有行的宽度
 * )
 * )
 * ),
 */
function down_personal_excel(array $data, $filename = '', array $mergeCells = array(), $style = array('height' => '20'), $sheetname = 'MCW', $objectPHPExcel = '', $i = 1, $end = true, $sheetIndex = 0)
{
//function downPersonalExcel(array $data, $style = array(), $sheetname= '', $filename='', $objectPHPExcel = '', $i = 1, $end = true, array $mergeCells=array(), $sheetIndex = 0) {
    $filename = !empty($filename) ? EscapeString($filename) : date('Y-m-dHis') . rand(1000, 999999);
    $filename = iconv("utf-8", 'gbk', $filename);
    if (!empty($data) && $data['fields'] && $filename) {
        if (empty($objectPHPExcel)) {
            vendor('phpexcel.PHPExcel');
            $objectPHPExcel = new \PHPExcel();
        }
        $column = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ');
        $rows = count($data['dataList']) + count($data['fields']);
        $cols = count($data['dataList'][0]) - 1;
        //创建人
        $objectPHPExcel->getProperties()->setCreator("MCW");
        //最后修改人
        $objectPHPExcel->getProperties()->setLastModifiedBy("MCW");
        // 标题
        $objectPHPExcel->getProperties()->setTitle('MCW' . date('Ymd'));
        //题目
        $objectPHPExcel->getProperties()->setCompany("MCW");
        //$objectPHPExcel->getActiveSheet()->getStyle(1)->getFont()->setBold(true);
        if (!empty($style['font'])) {
            //设置字体颜色等信息具体参数值参考官方文档
            $objectPHPExcel->getDefaultStyle()->getFont()->setName($style['font']);
        }

        //设置当前的sheet
        if ($sheetIndex > 0) {
            //可以通过如下的方式添加新的sheet暂未完善
            $objectPHPExcel->createSheet();
        }
        //操作当前sheet
        $objectPHPExcel->setActiveSheetIndex($sheetIndex);
        if ($sheetname) {
            //设置sheet的name
            $objectPHPExcel->getActiveSheet()->setTitle($sheetname);
        }
        //设置表格线
        if (empty($style['line'])) {
            $objectPHPExcel->getActiveSheet()->getStyle("A1:$column[$cols]" . ($rows + $i - 1))->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
            $objectPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getLeft()->getColor()->setARGB('FF993300');
            $objectPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getTop()->getColor()->setARGB('FF993300');
            $objectPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getBottom()->getColor()->setARGB('FF993300');
            $objectPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getRight()->getColor()->setARGB('FF993300');
        }

        foreach ($data['fields'] as $key => $val) {
            if (!empty($val)) {
                $temp_key = 0;
                if (!empty($style['height'])) {
                    $objectPHPExcel->getActiveSheet()->getRowDimension($i)->setRowHeight($style['height']);
                }
                foreach ($val as $k => $item) {
                    //标题字体颜色
                    if (!empty($item['color'])) {
                        $objectPHPExcel->getActiveSheet()->getStyle($column[$temp_key] . $i)->getFont()->getColor()->setRGB($item['color']);
                    }
                    //设置字体加粗
                    if (!empty($item['blod'])) {
                        $objectPHPExcel->getActiveSheet()->getStyle($column[$temp_key] . $i)->getFont()->setBold(true);
                    }
                    //设置表头的宽度
                    if (!empty($item['width'])) {
                        $objectPHPExcel->getActiveSheet()->getColumnDimension($column[$temp_key])->setWidth($item['width']);
                    }
                    if (!empty($item['height'])) {
                        $objectPHPExcel->getActiveSheet()->getRowDimension($i)->setRowHeight($item['height']);
                    }
                    //设置垂直居中
                    if (empty($item['remove_ver_align'])) {
                        $objectPHPExcel->getActiveSheet()->getStyle($column[$temp_key] . $i . ':' . $column[$cols] . ($rows + $i - 1))->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
                    }
                    //设置水平居中
                    if (empty($item['remove_text_align'])) {
                        $objectPHPExcel->getActiveSheet()->getStyle($column[$temp_key] . $i . ':' . $column[$cols] . ($rows + $i - 1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                    }
                    if (!empty($item['background'])) {
                        //标题背景颜色
                        $objectPHPExcel->getActiveSheet()->getStyle($column[$temp_key] . $i)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                        $objectPHPExcel->getActiveSheet()->getStyle($column[$temp_key] . $i)->getFill()->getStartColor()->setRGB($item['background']);
                    }
                    $objectPHPExcel->setActiveSheetIndex($sheetIndex)
                        ->setCellValueExplicit($column[$temp_key] . $i, $item['value'], PHPExcel_Cell_DataType::TYPE_STRING);
                    //如果需要合并单元格的话合并
                    if (!empty($item['merge'])) {
                        $objectPHPExcel->getActiveSheet()->mergeCells($column[$temp_key] . $i . ':' . $column[$temp_key + $item['merge']] . $i);
                        $temp_key += $item['merge'];
                    }
                    $temp_key++;
                }
                $i++;
            }
        }
        //$i = count($data['fields']) + 1;
        if (!empty($data['dataList'])) {
            foreach ($data['dataList'] as $key => $val) {  //每一行
                if (!empty($style['height'])) {
                    $objectPHPExcel->getActiveSheet()->getRowDimension($i)->setRowHeight($style['height']);
                }
                if ($val) {
                    $temp_k = 0;
                    foreach ($val as $v) {  //每一行里的每个单元格
                        $temp_v = $v;
                        //如果数组元素是数组的话，会进行单元格合并操作 add by v_plpeng 2017-04-10
                        /*
                        数组格式大概为
                        array(
                            0 => 'a',
                            1 => array(
                                num => '1',//向下合并单元格的个数
                                val => '单元格的值',//向下合并单元格的个数
                            ),
                            2 => 'c',
                        )
                        */
                        if (is_array($v) && isset($v['val']) && isset($v['num']) && $v['num'] > 0) {
                            $cell = $column[$temp_k] . $i . ':' . $column[$temp_k] . ($i + $v['num']);
                            $objectPHPExcel->getActiveSheet()->mergeCells($cell);
                            $temp_v = $v['val'];
                        }
                        if (is_int($temp_v)) {
                            $objectPHPExcel->setActiveSheetIndex($sheetIndex)
                                ->setCellValueExplicit($column[$temp_k] . $i, $temp_v, PHPExcel_Cell_DataType::TYPE_NUMERIC);
                        } else {
                            $objectPHPExcel->setActiveSheetIndex($sheetIndex)
                                ->setCellValueExplicit($column[$temp_k] . $i, $temp_v, PHPExcel_Cell_DataType::TYPE_STRING);
                        }
                        $temp_k++;
                    }
                    $i++;
                }
            }
            if ($mergeCells) {  //待合并的单元格
                foreach ($mergeCells as $cell) {
                    $objectPHPExcel->getActiveSheet()->mergeCells($cell);
                }
            }
        }
        if ($end == true) {
            $objectPHPExcel->setActiveSheetIndex(0);
            $ua = $_SERVER["HTTP_USER_AGENT"];
            if (preg_match("/MSIE/", $ua)) {
                $filename = urlencode($filename); //处理IE导出名称乱码
            }
//            ob_end_clean();//清除缓冲区,避免乱码
            //header('Content-Type: application/vnd.ms-excel');
            header("Content-Type: application/vnd.ms-excel; charset=UTF-8");
            header('Content-Disposition: attachment;filename="' . $filename . '.xls"');  //日期为文件名后缀
            header('Cache-Control: max-age=0');
            $objWriter = PHPExcel_IOFactory::createWriter($objectPHPExcel, 'Excel5');  //excel5为xls格式，excel2007为xlsx格式
            $objWriter->save('php://output');
        } else {
            return array('i' => $i, 'obj' => $objectPHPExcel);
        }
    }
}
