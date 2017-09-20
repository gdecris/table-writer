<?php


namespace Ivey\Writers;


use Illuminate\Support\Collection;

/**
 * @property Collection data
 * @property string output_string
 */
class DelimitedWriter extends AbstractFileWriter
{
    protected $delimiter = ',';

    protected $quote_values = false;

    protected $line_endings = "\n";

    protected $output_string = '';

    /**
     * @param $delimiter
     * @return $this
     */
    public function setDelimiter($delimiter)
    {
        $this->delimiter = $delimiter;

        return $this;
    }

    /**
     * @param bool $quote_values
     * @return $this
     */
    public function setQuoteValues(bool $quote_values)
    {
        $this->quote_values = $quote_values;

        return $this;
    }

    /**
     * @param $line_endings
     * @return $this
     */
    public function setLineEndings($line_endings)
    {
        $this->line_endings = $line_endings;

        return $this;
    }

    /**
     * @param null $filename
     */
    public function output($filename = null)
    {
        if (null !== $filename) {
            $filename = filter_var($filename, FILTER_SANITIZE_STRING, FILTER_FLAG_STRIP_LOW);
            header('Content-Type: text/csv');
            header('Content-Transfer-Encoding: binary');
            header("Content-Disposition: attachment; filename=\"$filename\"");
        }

        echo $this->generateData();
    }

    /**
     * @return string
     */
    private function generateData()
    {
        $this->output_string = '';

        $this->data->each(function ($row) {
            $this->output_string .=
                implode($this->delimiter, array_map([$this, 'formatValue'], $row)) .
                $this->line_endings;
        });

        return $this->output_string;
    }

    /**
     * @param $value
     * @return string
     */
    protected function formatValue($value)
    {
        if ( $this->quote_values ) {
            return "\"{$value}\"";
        }

        return $value;
    }
}