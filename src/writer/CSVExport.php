<?php

namespace data\writer;

class CSVExport
{
    /**
     * 表头字段
     * @var array
     */
    private $header;

    /**
     * 打开文件指针
     * @var resource
     */
    private $fp = null;

    /**
     * 根据文件获取指针
     * @param $file
     * @param string $mode 默认覆盖
     */
    public function __construct($file, $mode = 'w+')
    {
        /**
         * locale -a|grep zh_CN如果不支持查看系统编码
         * apt install language-pack-zh-hans
         */
        setlocale(LC_ALL, 'zh_CN.UTF-8');

        $path = pathinfo($file, PATHINFO_DIRNAME);
        $ext = pathinfo($file, PATHINFO_EXTENSION);
        /**
         * 目前支持csv格式
         */
        if ($ext == 'csv') {
            if (!is_dir($path)) {
                mkdir($path, 0775, true);
            }
            $this->fp = fopen($file, $mode);
            /**
             * UTF-8编码格式BOM头部
             */
            fwrite($this->fp, chr(0xEF) . chr(0xBB) . chr(0xBF));
        }
    }

    /**
     * 设置表头
     * @param $header
     */
    public function setHeader($header)
    {
        $this->header = $header;
        $content = implode(',', array_values($header)) . "\n";
        fwrite($this->fp, $content);
    }

    /**
     * 设置表内容
     * @param $data
     */
    public function setContent($data)
    {
        $content = '';
        foreach ($data as $row) {
            $tmp = [];
            foreach ($this->header as $key => $val) {
                if (isset($row[$key])) {
                    $tmp[] = $row[$key];
                } else {
                    $tmp[] = '';
                }

            }
            $content .= str_replace(["\n", "\r\n"], '', implode(',', $tmp)) . "\n";
        }
        fwrite($this->fp, $content);
    }

    public function __destruct()
    {
        fclose($this->fp);
    }
}