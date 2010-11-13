<?php
class ExcelComponent extends Object {
    
    var $errors = array();
    
    /**
     * readXls - Excelファイルを読み込む
     */
    function readXls($filepath, $colCount = null, $rowCount = null, $sheetIndex = null) 
    {
        //include the vendor class
        App::import('vendor','phpexcel/phpexcel');

        //ファイルを読み込む
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load($filepath);

        //シートオブジェクトの取得
        $sheets = array();
        if (is_null($sheetIndex)) {
            //すべて
            $sheets = $objPHPExcel->getAllSheets();
        } elseif (is_array($sheetIndex)) {
            foreach($sheetIndex as $idx) {
                $sheets[$idx] = $objPHPExcel->getSheet($idx);
            } 
        } elseif (is_int($sheetIndex)) {
            $sheets[$sheetIndex] = $objPHPExcel->getSheet($sheetIndex);
        }
        $data = array();
        if (empty($sheets)) {
            return $data;
        }

        //1シートごと処理
        foreach ($sheets as $s => $objSheet) {
            //シート名の取得
            $sheetTitle = $objSheet->getTitle();
            $data[$s]['title'] = $sheetTitle;

            //データ領域を確認
            $rowMax = $rowCount;
            if (is_null($rowCount)) {
                $rowMax = $objSheet->getHighestRow();
            }
            $colMax = $colCount;
            if (is_null($colCount)) {
                $colMax = $objSheet->getHighestColumn();
            }

            //1セルごとにテキストデータを取得
            $sheetData = array();
            for($r=1; $r<=$rowMax; $r++) { //rowは1はじまり
                 for($c=0; $c<=$colMax; $c++) { //colは0はじまり 0 = Aとなる
                     $objCell = $objSheet->getCellByColumnAndRow($c, $r);

                     $sheetData[$r][$c]= $this->_getText($objCell);
                 }
             }
             $data[$s]['data'] = $sheetData;
        }
        return $data;
        //include the vendor class
        App::import('vendor','phpexcel/phpexcel');

        //ファイルを読み込む
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load($filepath);

        //シートオブジェクトの取得
        $sheets = array();
        if (is_null($sheetIndex)) {
            //すべて
            $sheets = $objPHPExcel->getAllSheets();
        } elseif (is_array($sheetIndex)) {
            foreach($sheetIndex as $idx) {
                $sheets[$idx] = $objPHPExcel->getSheet($idx);
            } 
        } elseif (is_int($sheetIndex)) {
            $sheets[$sheetIndex] = $objPHPExcel->getSheet($sheetIndex);
        }
        $data = array();
        if (empty($sheets)) {
            return $data;
        }

        //1シートごと処理
        foreach ($sheets as $s => $objSheet) {
            //シート名の取得
            $sheetTitle = $objSheet->getTitle();
            $data[$s]['title'] = $sheetTitle;

            //データ領域を確認
            $rowMax = $rowCount;
            if (is_null($rowCount)) {
                $rowMax = $objSheet->getHighestRow();
            }
            $colMax = $colCount;
            if (is_null($colCount)) {
                $colMax = $objSheet->getHighestColumn();
            }

            //1セルごとにテキストデータを取得
            $sheetData = array();
            for($r=1; $r<=$rowMax; $r++) { //rowは1はじまり
                 for($c=0; $c<=$colMax; $c++) { //colは0はじまり 0 = Aとなる
                     $objCell = $objSheet->getCellByColumnAndRow($c, $r);

                     $sheetData[$r][$c]= $this->_getText($objCell);
                 }
             }
             $data[$s]['data'] = $sheetData;
        }
        return $data;
    }

    /**
     * write - Excelファイルに書き込む
     *
     * @param string $filepath テンプレートファイルのパス
     * @param array  $data
     * @param boolean $is_copy $filepathを テンプレートにしてコピーファイルを作るか
     */
    function writeXls($filepath, $data = array(), $is_copy = true)
    {
        //include the vendor class
        App::import('vendor','phpexcel/phpexcel');

        //ファイルを読み込む
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load($filepath);
        // set active sheet
        foreach ($data as $s => $rows) {
            $objPHPExcel->setActiveSheetIndex($s);
            $sheet = $objPHPExcel->getActiveSheet();
            
            foreach ($rows['data'] as  $r => $cols) {
                foreach ($cols as $c => $v) {
                    // update cell
                    $sheet->setCellValueByColumnAndRow($c, $r+1, $v);
                }
            }
        }
        if ($is_copy) {
            $target_filepath = TMP . "output_". date("YmdHis") .".xls"; //ファイル名生成
        } else {
            $target_filepath = $filepath;
        }

        //保存先のデータはパスとファイル名に分離が必要
        $target_basename = basename($target_filepath);
        $target_dir = dirname($target_filepath);

        // output excel file
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->setTempDir($target_dir);
        $objWriter->save($target_filepath);

        return $target_filepath;
    }

    /**
     * import - Excelファイルを読み込む
     */
    function import($file, $is_move_file = false, $destination = '', $type = "xls")
    {
    }


    /**
     * 指定したセルの文字列を取得する
     */
    function _getText($objCell = null)
    {
        if (is_null($objCell)) {
         return false;
        }

        $txtCell = "";

        //まずはgetValue()を実行
        $valueCell = $objCell->getValue();

        if (is_object($valueCell)) {
         //オブジェクトが返ってきたら、リッチテキスト要素を取得
         $rtfCell = $valueCell->getRichTextElements();
         //配列で返ってくるので、そこからさらに文字列を抽出
         $txtParts = array();
         foreach ($rtfCell as $v) {
            $txtParts[] = $v->getText();
         }
         //連結する
         $txtCell = implode("", $txtParts);

        } else {
         if (!empty($valueCell)) {
             $txtCell = $valueCell;
         }
        }

        return $txtCell;
    }
}
?>