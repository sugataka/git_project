class SampleController extends AppController
{
    var $components = array('Excel');
    
    /**
    * �_�E�����[�h
    */
    function download()
    {
        //�o�̓f�[�^�z��̐���
        $writeData = array(...);

        //Excel�ɏo��
        $return = $this->Excel->writeXls(WWW_ROOT . 'files/templates.xls', $writeData);

        //�G���[�`�F�b�N (�ȗ�)

        //�t�@�C�����ł���
        //�_�E�����[�h�J�n
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
    * �A�b�v���[�h
    */
    function upload()
    {

        // �ꊇ�A�b�v���[�h�t�@�C���m�F
        $tmp_path = TMP . 'upload_file.xls';
        $filename = $this->Excel->import($this->data['upload_file'], true, $tmp_path);
        if (!empty($this->Excel->errors)) {
            $this->Session->write('error_message', $this->Excel->errors);
            $this->redirect("/sample/index");
        }
        
        //�f�[�^�̓ǂݍ���
        $readData = $this->Excel->readXls($filename,11);

        //�ǂݍ��݌�̏��� (�ȗ�)
    }
}