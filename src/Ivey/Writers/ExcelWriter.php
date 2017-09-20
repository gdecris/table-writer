<?php


namespace Ivey\Writers;


/**
 * @property \PHPExcel excel
 * @property int row_count
 * @property  col_count
 */
class ExcelWriter extends AbstractFileWriter
{

    protected $defaults = [
        'creator' => 'Ivey Writer',
        'last_modified' => 'Ivey Writer',
        'title' => 'Ivey Writer Export',
        'subject' => '',
        'description' => '',
        'key_words' => '',
        'category' => '',
        'sheet_title' => 'Sheet 1',
        'accept_indexes' => false,
        'alt_on' => false,
        'highlight_on' => false,
        'col_format' => [],
        'file_name' => false,
        'auto_filter' => false,
        'value_callback' => [],
    ];

    protected $options = [];

    /**
     * @var \PHPExcel
     */
    protected $excel;

    protected $row_count;

    protected $col_count;

    protected $styles = [
        'default' => [],
        'alt' => [],
        'header' => []
    ];

    /**
     * @return mixed
     */
    public function getWriter()
    {
        return $this->excel;
    }

    public function autoFilter()
    {
        $this->options['auto_filter'] = true;

        return $this;
    }

    /**
     * Sets a style
     *
     * @param string $type  heading|default|alt
     * @param string $bgcolor   Hex string ex 000000
     * @param string $font_color    Hex string ex 000000
     * @param array $gradient   Array of hex strings ex [000000, FFFFFF]
     * @param bool $border
     * @return $this
     */
    public function style($type, $bgcolor, $font_color = '000000', $gradient = [], $border = true)
    {
        $this->styles[$type] = [
            'bgcolor' => $bgcolor,
            'font_color' => $font_color,
            'gradient' => $gradient,
            'border' => $border
        ];

        return $this;
    }

    /**
     * Sets a style
     *
     * @param string $bgcolor   Hex string ex 000000
     * @param string $font_color    Hex string ex 000000
     * @param array $gradient   Array of hex strings ex [000000, FFFFFF]
     * @param bool $border
     * @return $this
     */
    public function headerStyle($bgcolor, $font_color = '000000', $gradient = [], $border = true)
    {
        return $this->style('header', $bgcolor, $font_color, $gradient, $border);
    }

    /**
     * Sets a style
     *
     * @param string $bgcolor   Hex string ex 000000
     * @param string $font_color    Hex string ex 000000
     * @param array $gradient   Array of hex strings ex [000000, FFFFFF]
     * @param bool $border
     * @return $this
     */
    public function defaultStyle($bgcolor, $font_color = '000000', $gradient = [], $border = true)
    {
        return $this->style('default', $bgcolor, $font_color, $gradient, $border);
    }

    /**
     * Sets a style
     *
     * @param string $bgcolor   Hex string ex 000000
     * @param string $font_color    Hex string ex 000000
     * @param array $gradient   Array of hex strings ex [000000, FFFFFF]
     * @param bool $border
     * @return $this
     */
    public function altStyle($bgcolor, $font_color = '000000', $gradient = [], $border = true)
    {
        return $this->style('default', $bgcolor, $font_color, $gradient, $border);
    }


    /**
     * Sets up the Excel excel before output
     */
    protected function startWriter()
    {
        $options = array_merge($this->defaults, $this->options);

        $this->excel = new \PHPExcel();
        $this->excel->getProperties()
            ->setCreator($options['creator'])
            ->setLastModifiedBy($options['last_modified'])
            ->setTitle($options['title'])
            ->setSubject($options['subject'])
            ->setDescription($options['description'])
            ->setKeywords($options['key_words'])
            ->setCategory($options['category']);

        $this->excel->setActiveSheetIndex(0);
        $this->excel->getActiveSheet()
            ->setTitle( $options['sheet_title'] ?: 'Sheet 1' );
    }

    /**
     * @param null $filename
     */
    public function output($filename = null)
    {
        $this->startWriter();

        $filename = $filename ?? 'excel-export.xlsx';

        $this->col_count = count($this->data->first());
        $this->row_count = $this->data->count();

        $this->generateData();

        // Write the excel file
        $writer = \PHPExcel_IOFactory::createWriter($this->excel, 'Excel2007');
        $writer->save($temp_file = tempnam(sys_get_temp_dir(), 'excel-export-'));

        header('Pragma: hack');
        header('Vary: User-Agent');
        header("Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-disposition: attachment; filename=$filename");

        echo file_get_contents($temp_file);

        exit;
    }

    /**
     * Generates the excel sheet
     */
    private function generateData()
    {
        $this->data->each(function ($row, $row_index) {
            foreach ( $row as $col_index => $value ) {
                $this->addCell($value, $row_index+1, $col_index+1);
            }
            $this->applyStyles($row_index+1);
        });

        $this->applyAutoFilter();
    }

    /**
     * Adds a cell to te excel file
     *
     * @param $value
     * @param $row
     * @param $col
     * @return $this
     */
    public function addCell($value, $row, $col)
    {
        $col = $this->numberToColumn($col);

        if ( $value instanceof ExcelDataFormat ) {
            $format = $value->format;
            $value = $value->value;
            switch ( $format ) {
                case 'string':
                    $data_type = \PHPExcel_Cell_DataType::TYPE_STRING;
                    break;
                case 'formula':
                    $data_type = \PHPExcel_Cell_DataType::TYPE_FORMULA;
                    break;
                case 'date':
                    $skip_set = true;
                    $timestamp_value = strtotime($value);
                    $value = ( false === $timestamp_value ? $value : \PHPExcel_Shared_Date::PHPToExcel($timestamp_value) );
                    $this->excel->getActiveSheet()->setCellValue("{$col}{$row}", $value);
                    $this->excel->getActiveSheet()->getStyle("{$col}{$row}")
                        ->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_DATE_XLSX14);
                    break;
                case 'wrap':
                    $skip_set = true;
                    $this->excel->getActiveSheet()->setCellValue("{$col}{$row}", $value);
                    $this->excel->getActiveSheet()->getStyle("{$col}{$row}")->getAlignment()->setWrapText(true);
                    break;
                case 'url':
                    $skip_set = true;
                    $data_type = 'URL';
                    $this->excel->getActiveSheet()->setCellValue("{$col}{$row}", $value);
                    $this->excel->getActiveSheet()->getCell("{$col}{$row}")->getHyperlink()->setUrl($value);
                    break;
            }
        }

        if ( empty($data_type) ) {
            $this->excel->getActiveSheet()->setCellValue("{$col}{$row}", $value);
        } elseif ( !$skip_set ) {
            $this->excel->getActiveSheet()->setCellValueExplicit("{$col}{$row}", $value, $data_type);
        }

        return $this;
    }

    /**
     * @param $option_name
     * @return mixed
     */
    protected function option($option_name)
    {
        return array_get($this->options, $option_name);
    }

    /**
     * @param $row
     */
    private function applyStyles($row)
    {
        $col_start = $this->numberToColumn(1);
        $col_end = $this->numberToColumn($this->col_count);

        if ( $row == 1 && $header_style = $this->getHeaderStyle() ) {

            $style_info = $header_style;

        } elseif ( $style_info = $this->getDefaultStyle() ) {

            if ( $alt_row_style = $this->getAltStyle() ) {
                if ( $row % 2 == 0 ) {
                    $style_info = $alt_row_style;
                }
            }

        }

        if ( $style_info ) {
            $bgcolor = $style_info['bgcolor'];
            $color = $style_info['font_color'];
            $border = $style_info['border'];
            $gradient = $style_info['gradient'];

            $style = $this->getStyleArray($bgcolor, $color, $border, $gradient);

            $this->excel->getActiveSheet()
                ->getStyle("{$col_start}{$row}:{$col_end}{$row}")
                ->applyFromArray($style);
        }
    }

    /**
     *
     */
    public function applyAutoFilter()
    {
        //Turn on filters if options is set
        if ( $this->option('auto_filter') ) {
            $col_start = $this->numberToColumn(1);
            $col_end = $this->numberToColumn($this->col_count);

            $this->excel->getActiveSheet()->setAutoFilter("{$col_start}1:{$col_end}1");
            for ( $i = 1; $i < $this->col_count + 1; $i++ ) {
                $column = $this->numberToColumn($i);
                $this->excel->getActiveSheet()->getColumnDimension($column)->setAutoSize(true);
            }
        }
    }

    /**
     * @return mixed
     */
    public function getHeaderStyle()
    {
        return array_get($this->styles, 'header');
    }

    /**
     * @return mixed
     */
    public function getDefaultStyle()
    {
        return array_get($this->styles, 'default');
    }

    /**
     * @return mixed
     */
    public function getAltStyle()
    {
        return array_get($this->styles, 'alt');
    }


    /**
     * @param $bgcolor
     * @param $color
     * @param $border
     * @param bool $gradient
     * @return array
     */
    public function getStyleArray($bgcolor, $color, $border, $gradient = false) {
        $style = [];
        if ( is_array($gradient) && !empty($gradient) ) {
            $style['fill'] = [
                'type' => \PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startcolor' => ['argb' => $gradient[0]],
                'endcolor' => ['argb' => $gradient[1]]
            ];
        } elseif ( !empty($bgcolor) ) {
            $style['fill'] = [
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                'color' => ['argb' => $bgcolor]
            ];
        }
        if ( !empty($color) ) {
            $style['font'] = array(
                'color' => array('argb' => $color)
            );
        }
        if ( $border === true ) {
            $style['borders'] = [
                'allborders' => [
                    'style' => \PHPExcel_Style_Border::BORDER_THIN,
                    'color' => ['argb' => '666666']
                ]
            ];
        }
        return $style;
    }

    /**
     * @param $num
     * @param int $start
     * @param int $end
     * @return string
     */
    public function numberToColumn($num, $start = 65, $end = 90) {
        $sig = ($num < 0);
        $num = abs($num);
        $str = "";
        $cache = ($end - $start);
        while ( $num != 0 ) {
            $str = chr ( ( $num % $cache ) + $start - 1 ) . $str;
            $num = ( $num - ( $num % $cache) ) / $cache;
        }
        if ( $sig ) {
            $str = '-' . $str;
        }

        return $str;
    }

    /**
     * @param $value
     * @return ExcelDataFormat
     */
    public function formatDate($value)
    {
        return new ExcelDataFormat(ExcelDataFormat::DATE, $value);
    }

    /**
     * @param $value
     * @return ExcelDataFormat
     */
    public function formatWrap($value)
    {
        return new ExcelDataFormat(ExcelDataFormat::WRAP, $value);
    }

}