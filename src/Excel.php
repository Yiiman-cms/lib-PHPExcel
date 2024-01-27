<?php


namespace YiiMan\YiiLibExcel;

use PhpOffice\PhpSpreadsheet\Collection\Memory\SimpleCache1;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Settings;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
class Excel
{
    protected $excel;
    private $titles = [];
    private $Globalstyles = [];

    private $temp_path='';

    private $temp_file='';

    public function writer($type='Xlsx')
    {
        return IOFactory::createWriter($this->excel, IOFactory::WRITER_XLSX);
    }

    public function __construct()
    {
//        ignore_user_abort(true);
        ini_set('memory_limit', '-1');
        set_time_limit(0);
        ini_set('max_execution_time', 0);
        $this->temp_path=sys_get_temp_dir();
        $this->excel = new Spreadsheet();

        $this->excel->getProperties()->setCreator("Me")->setLastModifiedBy(
            "Me"
        )->setTitle(
            "My Excel Sheet"
        )->setSubject("My Excel Sheet")->setDescription("Excel Sheet")->setKeywords(
            "Excel Sheet"
        )->setCategory("Me");
        $this->excel->setActiveSheetIndex(0);
    }

    /**
     * Set active sheet (default is 0)
     * @param $index
     * @return $this
     */
    public function  set_active_sheet($index=0)
    {
        $this->excel->setActiveSheetIndex($index);
        return $this;
    }

    /**
     * Get active sheet
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function activeSheet():\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
    {
        return $this->excel->getActiveSheet();
    }

    /**
     * generate unique temp file path for write files
     * @return string
     */
    private function temp_file_path(): string
    {
        if (empty($this->temp_file)){
            $temp_file_path=tempnam($this->temp_path, 'prefix_'.uniqid().'_');
            $this->temp_file=$temp_file_path;
            return $temp_file_path;
        }else{
            return $this->temp_file;
        }
    }

    /**
     * get active sheet data as array
     * @return array
     */
    public function get_sheet_data():array
    {
        return $this->excel->getActiveSheet()->toArray();
    }

    /**
     * load excel file
     * @param $path
     * @return self
     */
    public function loadFile($path):self
    {
        /** Load $inputFileName to a PHPExcel Object  **/
        $this->excel = IOFactory::load($path);
        return $this;
    }

    /**
     * @param $coordinate
     * @return $this
     */
    public function freezeFirstRow($coordinate='A1'):self
    {
        $this->excel->getActiveSheet()->freezePane($coordinate);
        return $this;
    }

    /**
     * give cell and row number and return Name of cell num just like excel
     * @param $row_number
     * @param $cell_number
     * @return string
     */
    protected function cellNames($row_number, $cell_number):string
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
                $cache=new SimpleCache1();
                Settings::setCache($cache);
            }
            // </ Set Cache Mode >

            // < Set Language >
            {
                Settings::setLocale('fa_ir');
            }
            // </ Set Language >

            // < Orientation And Page Size >
            {
                $this->excel->getActiveSheet()->getPageSetup()->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
                $this->excel->getActiveSheet()->getPageSetup()->setPaperSize(PageSetup::PAPERSIZE_A4);
            }
            // </ Orientation And Page Size >
        }
        // </ Excel Settings >

    }

    /**
     * will Set Excel First Row As Header Titles
     *
     * @param array $titles ['title1','title2', ...]
     * @return self
     */
    public function setHeaders(array $titles):self
    {
        $this->titles = $titles;

        foreach ($this->titles as $tkey => $title) {
            $this->setFontname(1, $tkey, 'IRANSans');
            $this->setFontSize(1, $tkey, 11);
            $this->excel->getActiveSheet()
                ->setCellValue($this->cellNames(1, $tkey), $title);
        }

        $this->excel->getActiveSheet()->setAutoFilter('A1:'.$this->cellNames(1, count($this->titles)-1));
        return $this;
    }

    /**
     * Will Set Array of Styles For Cells and Title Rows
     * <code>
     * $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray(
     * [
     * 'font' => [
     *      'name' => 'Arial',
     *      'bold' => true,
     *      'italic' => false,
     *      'underline' => Font::UNDERLINE_DOUBLE,
     *      'strikethrough' => false,
     * 'color' => [
     *       'rgb' => '808080'
     * ]
     * ],
     * 'borders' => [
     *      'bottom' => [
     *          'borderStyle' => Border::BORDER_DASHDOT,
     *           'color' => [
     *                'rgb' => '808080'
     *          ]
     *      ],
     *      'top' => [
     *            'borderStyle' => Border::BORDER_DASHDOT,
     *            'color' => [
     *                'rgb' => '808080'
     *               ]
     *      ]
     * ],
     * 'alignment' => [
     *          'horizontal' => Alignment::HORIZONTAL_CENTER,
     *          'vertical' => Alignment::VERTICAL_CENTER,
     *          'wrapText' => true,
     *      ],
     *      'quotePrefix'    => true
     * ]
     * );
     * </code>
     * @param array $style_array
     * @return self
     */
    public function setGlobalStyles(array $style_array):self
    {
        if (!empty($style_array)) {
            $this->Globalstyles = $style_array;
            foreach ($this->titles as $tkey => $title) {
                $this->excel->getActiveSheet()->getStyle($this->cellNames(1, $tkey))->applyFromArray($style_array);
            }
        }
        return $this;
    }

    /**
     * set style on single cell
     * <code>
     * $spreadsheet->getActiveSheet()->getStyle('B2')->applyFromArray(
     * [
     * 'font' => [
     *      'name' => 'Arial',
     *      'bold' => true,
     *      'italic' => false,
     *      'underline' => Font::UNDERLINE_DOUBLE,
     *      'strikethrough' => false,
     * 'color' => [
     *       'rgb' => '808080'
     * ]
     * ],
     * 'borders' => [
     *      'bottom' => [
     *          'borderStyle' => Border::BORDER_DASHDOT,
     *           'color' => [
     *                'rgb' => '808080'
     *          ]
     *      ],
     *      'top' => [
     *            'borderStyle' => Border::BORDER_DASHDOT,
     *            'color' => [
     *                'rgb' => '808080'
     *               ]
     *      ]
     * ],
     * 'alignment' => [
     *          'horizontal' => Alignment::HORIZONTAL_CENTER,
     *          'vertical' => Alignment::VERTICAL_CENTER,
     *          'wrapText' => true,
     *      ],
     *      'quotePrefix'    => true
     * ]
     * );
     * </code>
     *
     * @param int $row_number
     * @param int $column_number
     * @param array $style
     *

     * @return self
     */
    public function setStyle(int $row_number, int $column_number, array $style):self
    {
        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->applyFromArray($style);
        return $this;
    }

    /**
     * set a Cell Hyperlink
     * @param int $row_number
     * @param int $column_number
     * @param string $url
     * @return self
     */
    public function setHyperLink(int $row_number, int $column_number, string $url):self
    {
        $this->excel->getActiveSheet()->setCellValue($this->cellNames($row_number, $column_number), $url);
        $this->excel->getActiveSheet()->getCell($this->cellNames($row_number, $column_number))->getHyperlink()->setUrl($url);
        return $this;
    }

    /**
     * format a cell as separated number
     * @param int $row_number
     * @param int $column_number
     * @return self
     */
    public function setFormatNumberCell(int $row_number, int $column_number):self
    {
        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->getNumberFormat()
            ->setFormatCode('#,##0');
        return $this;
    }

    /**
     * set Cell Text Color
     * @param int $row_number
     * @param int $column_number
     * @param string $color like FF000000 ,You can use Color::Color_constants
     * @return self
     */
    public function setTextColor(int $row_number, int $column_number, string $color):self
    {
        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['color' => ['rgb' => $color]]]);
        return $this;
    }

    /**
     * set Cell Text Bold
     * @param int $row_number
     * @param int $column_number
     * @return self
     */
    public function setTextBold(int $row_number, int $column_number):self
    {
        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['bold' => true]]);
        return $this;
    }

    /**
     * set Cell Text Italic
     * @param int $row_number
     * @param int $column_number
     * @return self
     */
    public function setTextItalic(int $row_number, int $column_number):self
    {
        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['italic' => true]]);
        return $this;
    }

    /**
     * set Cell Text Underline
     * @param int $row_number
     * @param int $column_number
     * @return self
     */
    public function setTextUnderline(int $row_number, int $column_number):self
    {
        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['underline' => true]]);
        return $this;
    }


    /**
     * this function will lock your cell for prevent from edit
     * @param int $row_number
     * @param int $column_number
     * @return self
     */
    public function setLockCell(int $row_number, int $column_number):self
    {
        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->getProtection()->setLocked(
            Protection::PROTECTION_PROTECTED
        );
        return $this;
    }

    /**
     * set cell fill color in hex(rgb) mode
     * @param int $row_number
     * @param int $column_number
     * @param string $color like FF000000 ,You can use Color::Color_constants
     * @return self
     */
    public function setFillColor(int $row_number, int $column_number, string $color):self
    {

        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['fill' => ['color' => ['rgb' => $color], 'type' => Fill::FILL_SOLID]]);
        return $this;
    }

    /**
     * this function will hide array of columns
     * @param array $columns_names
     * @return self
     */
    public function hideColumns(array $columns_names):self
    {
        foreach ($columns_names as $column_name) {
            $this->excel->getActiveSheet()->getColumnDimension($column_name)->setVisible(false);
        }
        return $this;
    }

    /**
     * Set excel Sheet as RTL
     * @return $this
     */
    public function isRTL():self
    {
        $this->excel->getActiveSheet()
            ->setRightToLeft(true);
        return $this;
    }

    /**
     * will set colored text in one cell
     * @param int $row_number
     * @param int $column_number
     * @param string $text
     * @param string $color like FF000000 ,You can use Color::Color_constants
     * @return self
     */
    public function RichText(int $row_number, int $column_number, string $text, string $color):self
    {
        $objRichText = new RichText();
//        $objRichText->createText('This invoice is ');

        $objPayable = $objRichText->createTextRun($text);
        $objPayable->getFont()->setBold(true);
        $objPayable->getFont()->setItalic(true);
        $objPayable->getFont()->setColor(new Color($color));

//        $objRichText->createText(', unless specified otherwise on the invoice.');

        $this->excel->getActiveSheet()->getCell($this->cellNames($row_number, $column_number))->setValue($objRichText);
        return $this;
    }

    /**
     * @param array $columns
     * @return self
     */
    public function setColumnsWidth(array $columns):self
    {
        foreach ($columns as $columnName => $width) {
            if (is_bool($width)) {
                $this->excel->getActiveSheet()->getColumnDimension($columnName)->setAutoSize(true);
            } else {
                $this->excel->getActiveSheet()->getColumnDimension($columnName)->setWidth($width);
            }
        }
        return $this;
    }

    /**
     * Set a value on activated sheet cell
     * @param int $row_number
     * @param int $column_number
     * @param $value
     * @return $this
     */
    public function setValue(int $row_number, int $column_number, $value):self
    {
        $this->setFontname($row_number, $column_number, 'IRANSans');
        $this->setFontSize($row_number, $column_number, 11);
        $this->excel->getActiveSheet()->setCellValue(
            $this->cellNames($row_number, $column_number),
            $value
        );
        if (!empty($this->Globalstyles)) {
            $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->applyFromArray($this->Globalstyles);
        }
        return $this;
    }

    /**
     * change font name for a cell
     * @param $row_number
     * @param $column_number
     * @param $fontName
     * @return self
     */
    public function setFontname($row_number,$column_number,$fontName):self
    {
        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['name' => 'IRANSans']]);
        return $this;
    }

    /**
     * change font size
     * @param $row_number
     * @param $column_number
     * @param $font_size
     * @return self
     */
    public function setFontSize($row_number,$column_number,$font_size=11):self{
        $this->excel->getActiveSheet()->getStyle($this->cellNames($row_number, $column_number))->applyFromArray(['font' => ['size' => $font_size]]);
        return $this;
    }

    /**
     * Write latest changes and get file path
     *
     * Please delete file after do some things
     *
     * @return string
     */
    public function write_and_get_file_path():string
    {
        $this->writer()->save(
            $this->temp_file_path(). '.xls'
        );
        return $this->temp_file_path(). '.xls';
    }


    /**
     * Write changes on temp file and start download on client browser
     * @param $file_name
     * @param $title
     * @return mixed
     */
    public function download($file_name = 'excel.xls', $title = 'excel file')
    {
        
        $this->excel->getActiveSheet()->setTitle($title);
//										PHPExcel_Settings::setZipClass(PHPExcel_Settings::PCLZIP);
        $this->writer()->save(
            $this->temp_file_path(). '.xls'
        );

        $file = $this->temp_file_path() . '.xls';

        header("Content-Description: File Transfer");
        header("Content-Type: application/octet-stream");
        header("Content-Disposition: attachment; filename=\"" . basename($file_name) . "\"");

        readfile($file);
        unset($file);
        exit();
    }
}
