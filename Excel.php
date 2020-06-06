<?php


namespace system\lib;


class Excel
{
    protected $excel;
    private $titles = [];
    private $Globalstyles = [];
    private $style = [];
    private $writer;
    private $data;
    private $activeSheet;

    public function __construct()
    {
//        ignore_user_abort(true);
        ini_set('memory_limit', '-1');
        set_time_limit(0);
        ini_set('max_execution_time', 0);
        include_once __DIR__ . '/../../crm_include/lib/PHPExcel.php';
        $this->excel = $objPHPExcel = new \PHPExcel();

        $this->writer = \PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');
        $this->excel->getProperties()->setCreator("Me")->setLastModifiedBy(
            "Me"
        )->setTitle(
            "My Excel Sheet"
        )->setSubject("My Excel Sheet")->setDescription("Excel Sheet")->setKeywords(
            "Excel Sheet"
        )->setCategory("Me");
        $this->excel->setActiveSheetIndex(0);
        $this->activeSheet = $this->excel->getActiveSheet();
    }

    public function loadFile($post_field_name)
    {
        if (!empty($_FILES[$post_field_name])) {
           $move= move_uploaded_file($_FILES[$post_field_name]['tmp_name'], __DIR__ . '/excel.xls');

        }
        /** Load $inputFileName to a PHPExcel Object  **/
        $this->excel = \PHPExcel_IOFactory::load(__DIR__ . '/excel.xls');
//        unlink(__DIR__ . '/excel.xls');
        return $this->excel->getActiveSheet()->toArray();
    }

    public function freezFirstRow()
    {
        $this->activeSheet->freezePaneByColumnAndRow();
    }

    /**
     * give cell and row number and return charachtred cell num just like excel
     * @param $row_number
     * @param $cell_number
     * @return string
     */
    protected function cellNames($row_number, $cell_number)
    {
        $char_array =
            [
                'A',
                'B',
                'C',
                'D',
                'E',
                'F',
                'G',
                'H',
                'I',
                'J',
                'K',
                'L',
                'M',
                'N',
                'O',
                'P',
                'Q',
                'R',
                'S',
                'T',
                'U',
                'V',
                'W',
                'X',
                'Y',
                'Z'
            ];
        if ($cell_number <= 26) {
            return $char_array[$cell_number] . $row_number;
        }
    }

    protected function LoadSettings()
    {

        // < Excel Settings >
        {
            // < Set Cache Mode >
            {
                $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
                \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, ['memoryCacheSize' => '8MB']);
            }
            // </ Set Cache Mode >

            // < Set Language >
            {
                \PHPExcel_Settings::setLocale('fa_ir');
            }
            // </ Set Language >

            // < Orientation And Page Size >
            {
                $this->activeSheet->getPageSetup()->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                $this->activeSheet->getPageSetup()->setPaperSize(\PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
            }
            // </ Orientation And Page Size >
        }
        // </ Excel Settings >

    }

    /**
     * will Set Excel First Row As Header Titles
     *
     * @param array $titles ['title1','title2', ...]
     */
    public function setHeaders(array $titles)
    {
        $this->titles = $titles;

        foreach ($this->titles as $tkey => $title) {
            $this->setFontname(1, $tkey, 'IRANSans');
            $this->setFontSize(1, $tkey, 11);
            $this->activeSheet
                ->setCellValue($this->cellNames(1, $tkey), $title);
        }

        $this->excel->getActiveSheet()->setAutoFilter('A1:'.$this->cellNames(1, count($this->titles)-1));
    }

    /**
     * Will Set Style For Cells and Title Rows
     * @param array $style_array
     * @throws \PHPExcel_Exception
     */
    public function setGlobalStyles(array $style_array)
    {
        if (!empty($style_array)) {
            $this->Globalstyles = $style_array;
            foreach ($this->titles as $tkey => $title) {
                $this->activeSheet->getStyle($this->cellNames(1, $tkey))->applyFromArray($style_array);
            }
        }
    }

    /**
     * set style on single cell
     * @param int $row_number
     * @param int $column_number
     * @param array $style
     * @throws \PHPExcel_Exception
     */
    public function setStyle(int $row_number, int $column_number, array $style)
    {
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->applyFromArray($style);

    }

    /**
     * set an Cell Hyperlink
     * @param int $row_number
     * @param int $column_number
     * @param string $url
     * @throws \PHPExcel_Exception
     */
    public function setHyperLink(int $row_number, int $column_number, string $url)
    {
        $this->activeSheet->setCellValue($this->cellNames($row_number, $column_number), $url);
        $this->activeSheet->getCell($this->cellNames($row_number, $column_number))->getHyperlink()->setUrl($url);

    }

    /**
     * format an cell as separated number
     * @param int $row_number
     * @param int $column_number
     * @throws \PHPExcel_Exception
     */
    public function setFormatNumberCell(int $row_number, int $column_number)
    {
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->getNumberFormat()
            ->setFormatCode('#,##0');

    }

    /**
     * set Cell Text Color
     * @param int $row_number
     * @param int $column_number
     * @param string $color color in hex without #
     * @throws \PHPExcel_Exception
     */
    public function setTextColor(int $row_number, int $column_number, string $color)
    {
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['color' => ['rgb' => $color]]]);
    }

    /**
     * set Cell Text Bold
     * @param int $row_number
     * @param int $column_number
     * @throws \PHPExcel_Exception
     */
    public function setTextBold(int $row_number, int $column_number)
    {
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['bold' => true]]);
    }

    /**
     * set Cell Text Italic
     * @param int $row_number
     * @param int $column_number
     * @throws \PHPExcel_Exception
     */
    public function setTextItalic(int $row_number, int $column_number)
    {
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['italic' => true]]);
    }

    /**
     * set Cell Text Underline
     * @param int $row_number
     * @param int $column_number
     * @throws \PHPExcel_Exception
     */
    public function setTextUnderline(int $row_number, int $column_number)
    {
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['underline' => true]]);
    }


    /**
     * this function will lock your cell for prevent from edit
     * @param int $row_number
     * @param int $column_number
     */
    public function setLockCell(int $row_number, int $column_number)
    {
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->getProtection()->setLocked(
            \PHPExcel_Style_Protection::PROTECTION_PROTECTED
        );
    }

    /**
     * set cell fill color in hex(rgb) mode
     * @param int $row_number
     * @param int $column_number
     * @param string $color hex color without #
     * @throws \PHPExcel_Exception
     */
    public function setFillColor(int $row_number, int $column_number, string $color)
    {
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['fill' => ['color' => ['rgb' => $color], 'type' => \PHPExcel_Style_Fill::FILL_SOLID]]);

    }

    /**
     * this function will hide array of columns
     * @param array $columns_name
     * @throws \PHPExcel_Exception
     */
    public function hideColumns(array $columns_names)
    {
        foreach ($columns_names as $column_name) {
            $this->activeSheet->getColumnDimension($column_name)->setVisible(false);
        }
    }

    public function isRTL()
    {
        $this->activeSheet
            ->setRightToLeft(true);

    }

    /**
     * will set colored text in one cell
     * @param int $row_number
     * @param int $column_number
     * @param string $text
     * @param string $color color in hex(ARGB) format without #
     * @throws \PHPExcel_Exception
     */
    public function RichText(int $row_number, int $column_number, string $text, string $color)
    {
        $objRichText = new \PHPExcel_RichText();
//        $objRichText->createText('This invoice is ');

        $objPayable = $objRichText->createTextRun($text);
        $objPayable->getFont()->setBold(true);
        $objPayable->getFont()->setItalic(true);
        $objPayable->getFont()->setColor(new \PHPExcel_Style_Color($color));

//        $objRichText->createText(', unless specified otherwise on the invoice.');

        $this->activeSheet->getCell($this->cellNames($row_number, $column_number))->setValue($objRichText);

    }

    /**
     * @param array $columns
     * @throws \PHPExcel_Exception
     */
    public function setColumnsWidth(array $columns)
    {
        foreach ($columns as $columnName => $width) {
            if (is_bool($width)) {
                $this->activeSheet->getColumnDimension($columnName)->setAutoSize(true);
            } else {
                $this->activeSheet->getColumnDimension($columnName)->setWidth($width);
            }
        }
    }

    public function setValue(int $row_number, int $column_number, $value)
    {
        $this->setFontname($row_number, $column_number, 'IRANSans');
        $this->setFontSize($row_number, $column_number, 11);
        $this->activeSheet->setCellValue(
            $this->cellNames($row_number, $column_number),
            $value
        );
        if (!empty($this->Globalstyles)) {
            $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->applyFromArray($this->Globalstyles);
        }
    }

    /**
     * change font name for a cell
     * @param $row_number
     * @param $column_number
     * @param $fontName
     * @throws \PHPExcel_Exception
     */
    public function setFontname($row_number,$column_number,$fontName)
    {
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['name' => 'IRANSans']]);
    }

    /**
     * change font size
     * @param $row_number
     * @param $column_number
     * @param $font_size
     * @throws \PHPExcel_Exception
     */
    public function setFontSize($row_number,$column_number,$font_size){
        $this->activeSheet->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['size' => 11]]);
    }

    public function execute($file_name = 'excel', $title = 'excel file')
    {

        $this->activeSheet->setTitle($title);
//										PHPExcel_Settings::setZipClass(PHPExcel_Settings::PCLZIP);
        $this->writer->save(
            __DIR__ . '/' . $file_name . '.xls'
        );

        $file = __DIR__ . '/' . $file_name . '.xls';

        header("Content-Description: File Transfer");
        header("Content-Type: application/octet-stream");
        header("Content-Disposition: attachment; filename=\"" . basename($file) . "\"");

        readfile($file);
        exit();
    }
}
