<?php


namespace Ivey\Writers;


/**
 * @property  value
 * @property  format
 */
class ExcelDataFormat
{
    public $format;
    public $value;

    const DATE = 'date';
    const WRAP = 'wrap';
    const STRING = 'string';
    const FORMULA = 'formula';
    const URL = 'url';


    public function __construct($format, $value)
    {
        $this->format = $format;
        $this->value = $value;
    }

}