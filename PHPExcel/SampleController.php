class SampleController extends AppController
{
    var $components = array('Excel');
    
    /**
    * ダウンロード
    */
    function download()
    {
        //出力データ配列の生成
        $writeData = array(...);

        //Excelに出力
        $return = $this->Excel->writeXls(WWW_ROOT . 'files/templates.xls', $writeData);

        //エラーチェック (省略)

        //ファイルができた
        //ダウンロード開始
        $media_id = basename($return);
        $media_name = substr($media_id, 0, strlen($media_id) - 4);

        $this->view ='media';
        $params = array(
            'id' => $media_id,
            'name' => $media_name,
            'download' => true,
            'extension' => 'xls',
            'path' => dirname($return) . DS,
            'mime' => "application/vnd.ms-excel"
            );

        $this->set($params);
    }


    /**
    * アップロード
    */
    function upload()
    {

        // 一括アップロードファイル確認
        $tmp_path = TMP . 'upload_file.xls';
        $filename = $this->Excel->import($this->data['upload_file'], true, $tmp_path);
        if (!empty($this->Excel->errors)) {
            $this->Session->write('error_message', $this->Excel->errors);
            $this->redirect("/sample/index");
        }
        
        //データの読み込み
        $readData = $this->Excel->readXls($filename,11);

        //読み込み後の処理 (省略)
    }
}