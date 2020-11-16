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
     * @param $header
     */
    public function setHeader($header)
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
  <Title>Untitled Spreadsheet</Title>
  <Author>Unknown Creator</Author>
  <LastAuthor>Microsoft Office User</LastAuthor>
  <Created>{$date}T{$time}Z</Created>
  <LastSaved>{$date}T{$time}Z</LastSaved>
  <Company>Microsoft Corporation</Company>
  <Version>16.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
 </Styles>
here;

        fwrite($this->fp, $content);
    }

    /**
     * 设置sheet头
     * @param $name
     * @param $header
     */
    public function setSheetHeader($name, $header)
    {
        $content = <<<here
 <Worksheet ss:Name="$name">
  <Table>
here;
        $this->header = $header;
        $content .= "<Row>\n";
        foreach ($header as $val) {
            $content .= <<<here
<Cell><Data ss:Type="String">$val</Data></Cell>
here;
        }
        $content .= "</Row>\n";
        fwrite($this->fp, $content);

    }

    /**
     * 设置sheet表内容
     * @param $data
     */
    public function setContent($data)
    {
        $content = '';
        foreach ($data as $row) {
            $content .= "<Row>\n";
            foreach ($this->header as $key => $val) {
                if (!isset($row[$key])) { // 防止没有数据情况
                    $content .= <<<here
<Cell><Data ss:Type="String"></Data></Cell>
here;
                    continue;
                }
                $content .= <<<here
<Cell><Data ss:Type="String">$row[$key]</Data></Cell>
here;
            }
            $content .= "</Row>\n";
        }
        fwrite($this->fp, $content);
    }

    public function setSheetFooter()
    {
        $content = <<<here
        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Print>
    <ValidPrinterInfo/>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
   <AllowFormatCells/>
   <AllowSizeCols/>
   <AllowSizeRows/>
   <AllowInsertCols/>
   <AllowInsertRows/>
   <AllowInsertHyperlinks/>
   <AllowDeleteCols/>
   <AllowDeleteRows/>
   <AllowSort/>
   <AllowFilter/>
   <AllowUsePivotTables/>
  </WorksheetOptions>
 </Worksheet>
here;
        fwrite($this->fp, $content);
    }

    /**
     * 设置结尾
     */
    public function setFooter()
    {
        $content = <<<here
        </Workbook>
here;
        fwrite($this->fp, $content);
    }

    public function __destruct()
    {
        fclose($this->fp);
    }
}