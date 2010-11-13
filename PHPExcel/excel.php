<?php
class ExcelComponent extends Object {
    
    var $errors = array();
    
    /**
     * readXls - Excel�t�@�C����ǂݍ���
     */
    function readXls($filepath, $colCount = null, $rowCount = null, $sheetIndex = null) 
    {
        //include the vendor class
        App::import('vendor','phpexcel/phpexcel');

        //�t�@�C����ǂݍ���
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load($filepath);

        //�V�[�g�I�u�W�F�N�g�̎擾
        $sheets = array();
        if (is_null($sheetIndex)) {
            //���ׂ�
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

        //1�V�[�g���Ə���
        foreach ($sheets as $s => $objSheet) {
            //�V�[�g���̎擾
            $sheetTitle = $objSheet->getTitle();
            $data[$s]['title'] = $sheetTitle;

            //�f�[�^�̈���m�F
            $rowMax = $rowCount;
            if (is_null($rowCount)) {
                $rowMax = $objSheet->getHighestRow();
            }
            $colMax = $colCount;
            if (is_null($colCount)) {
                $colMax = $objSheet->getHighestColumn();
            }

            //1�Z�����ƂɃe�L�X�g�f�[�^���擾
            $sheetData = array();
            for($r=1; $r<=$rowMax; $r++) { //row��1�͂��܂�
                 for($c=0; $c<=$colMax; $c++) { //col��0�͂��܂� 0 = A�ƂȂ�
                     $objCell = $objSheet->getCellByColumnAndRow($c, $r);

                     $sheetData[$r][$c]= $this->_getText($objCell);
                 }
             }
             $data[$s]['data'] = $sheetData;
        }
        return $data;
        //include the vendor class
        App::import('vendor','phpexcel/phpexcel');

        //�t�@�C����ǂݍ���
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $objPHPExcel = $objReader->load($filepath);

        //�V�[�g�I�u�W�F�N�g�̎擾
        $sheets = array();
        if (is_null($sheetIndex)) {
            //���ׂ�
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

        //1�V�[�g���Ə���
        foreach ($sheets as $s => $objSheet) {
            //�V�[�g���̎擾
            $sheetTitle = $objSheet->getTitle();
            $data[$s]['title'] = $sheetTitle;

            //�f�[�^�̈���m�F
            $rowMax = $rowCount;
            if (is_null($rowCount)) {
                $rowMax = $objSheet->getHighestRow();
            }
            $colMax = $colCount;
            if (is_null($colCount)) {
                $colMax = $objSheet->getHighestColumn();
            }

            //1�Z�����ƂɃe�L�X�g�f�[�^���擾
            $sheetData = array();
            for($r=1; $r<=$rowMax; $r++) { //row��1�͂��܂�
                 for($c=0; $c<=$colMax; $c++) { //col��0�͂��܂� 0 = A�ƂȂ�
                     $objCell = $objSheet->getCellByColumnAndRow($c, $r);

                     $sheetData[$r][$c]= $this->_getText($objCell);
                 }
             }
             $data[$s]['data'] = $sheetData;
        }
        return $data;
    }

    /**
     * write - Excel�t�@�C���ɏ�������
     *
     * @param string $filepath �e���v���[�g�t�@�C���̃p�X
     * @param array  $data
     * @param boolean $is_copy $filepath�� �e���v���[�g�ɂ��ăR�s�[�t�@�C������邩
     */
    function writeXls($filepath, $data = array(), $is_copy = true)
    {
        //include the vendor class
        App::import('vendor','phpexcel/phpexcel');

        //�t�@�C����ǂݍ���
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
            $target_filepath = TMP . "output_". date("YmdHis") .".xls"; //�t�@�C��������
        } else {
            $target_filepath = $filepath;
        }

        //�ۑ���̃f�[�^�̓p�X�ƃt�@�C�����ɕ������K�v
        $target_basename = basename($target_filepath);
        $target_dir = dirname($target_filepath);

        // output excel file
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->setTempDir($target_dir);
        $objWriter->save($target_filepath);

        return $target_filepath;
    }

    /**
     * import - Excel�t�@�C����ǂݍ���
     */
    function import($file, $is_move_file = false, $destination = '', $type = "xls")
    {
    }


    /**
     * �w�肵���Z���̕�������擾����
     */
    function _getText($objCell = null)
    {
        if (is_null($objCell)) {
         return false;
        }

        $txtCell = "";

        //�܂���getValue()�����s
        $valueCell = $objCell->getValue();

        if (is_object($valueCell)) {
         //�I�u�W�F�N�g���Ԃ��Ă�����A���b�`�e�L�X�g�v�f���擾
         $rtfCell = $valueCell->getRichTextElements();
         //�z��ŕԂ��Ă���̂ŁA�������炳��ɕ�����𒊏o
         $txtParts = array();
         foreach ($rtfCell as $v) {
            $txtParts[] = $v->getText();
         }
         //�A������
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