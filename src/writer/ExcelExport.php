<?php

namespace data\writer;

class ExcelExport
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
         * 目前支持xls格式
         */
        if ($ext == 'xls') {
            if (!is_dir($path)) {
                mkdir($path, 0775, true);
            }
            $this->fp = fopen($file, $mode);
        }

    }

    /**
     * 设置表头
     * @param $header
     */
    public function setHeader($header)
    {
        $content = <<<here
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
	<meta http-equiv="Content-type" content="text/html;charset=UTF-8" />
	<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>excelName</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
</head>
<body>
	<table>
        <thead>
here;

        $this->header = $header;
        $content .= "<tr>\n";
        foreach ($header as $val) {
            $content .= "<td>$val</td>\n";
        }
        $content .= "</tr>\n";
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
            $content .= "<tr>\n";
            foreach ($this->header as $key => $val) {
                if (!isset($row[$key])) { // 防止没有数据情况
                    $content .= "<td></td>\n";
                    continue;
                }
                if (isset($row[$key]['style'])) {
                    $content .= "<td {$row[$key]['style']}>{$row[$key]['value']}</td>\n";
                } else {
                    $tmp = $row[$key];
                    if (is_numeric($tmp) && !empty($tmp)) {
                        $content .= "<td x:str >$tmp</td>\n";
                    } else {
                        $content .= "<td>$tmp</td>\n";
                    }

                }
            }
            $content .= "</tr>\n";
        }
        fwrite($this->fp, $content);
    }

    /**
     * 设置结尾
     */
    public function setFooter()
    {
        $content = <<<here
        </tbody>
    </table>
</body>
</html>
here;
        fwrite($this->fp, $content);
    }

    public function __destruct()
    {
        fclose($this->fp);
    }
}