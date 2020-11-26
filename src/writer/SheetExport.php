<?php

namespace data\writer;

class SheetExport
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
     */
    public function setHeader()
    {
        $date = date('Y-m-d');
        $time = date('H:i:s');
        $content = <<<here
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Microsoft Office User</Author>
  <LastAuthor>Microsoft Office User</LastAuthor>
  <Created>{$date}T{$time}Z</Created>
  <LastSaved>{$date}T{$time}Z</LastSaved>
  <Version>16.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <ActiveSheet>1</ActiveSheet>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="等线" x:CharSet="134" ss:Size="12" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
 </Styles>\n
here;

        fwrite($this->fp, $content);
    }

    /**
     * 设置sheet头部
     * @param $name
     * @param $header
     */
    public function setSheetHead($name = 'Sheet1')
    {
        $content = <<<here
 <Worksheet ss:Name="$name">
  <Table x:FullColumns="1" x:FullRows="1" ss:DefaultColumnWidth="65" ss:DefaultRowHeight="16">\n
here;
        fwrite($this->fp, $content);
    }

    /**
     * 设置sheet表行
     * @param $data
     */
    public function setRow($data)
    {
        $content = '';
        foreach ($data as $row) {
            $content .= "   <Row>\n";
            foreach ($row as $cell) {
                $priverty = '';// 单元格属性
                if (isset($cell['index'])) { // 设置开始索引
                    $priverty .= <<<here
ss:Index="{$cell['index']}"
here;
                }
                if (isset($cell['across']) && $cell['across'] > 0) { // 向右合并单元格
                    $priverty .= <<<here
 ss:MergeAcross="{$cell['across']}"
here;
                }
                if (isset($cell['down']) && $cell['down'] > 0) { // 向下合并单元格
                    $priverty .= <<<here
 ss:MergeDown="{$cell['down']}"
here;
                }
                $type = 'String';
                if (isset($cell['type'])) { // 设置单元格数据类型
                    $type = $cell['type'];
                }
                $content .= <<<here
    <Cell $priverty><Data ss:Type="$type">{$cell['value']}</Data></Cell>\n
here;
            }
            $content .= "   </Row>\n";
        }
        fwrite($this->fp, $content);
    }

    public function setSheetFooter()
    {
        $content = <<<here
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>\n
here;
        fwrite($this->fp, $content);
    }

    /**
     * 设置结尾
     */
    public function setFooter()
    {
        $content = <<<here
</Workbook>\n
here;
        fwrite($this->fp, $content);
    }

    public function __destruct()
    {
        fclose($this->fp);
    }
}