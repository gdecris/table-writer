<?php


use Ivey\Writers\DelimitedWriter;

class DelimitedFileTest extends PHPUnit_Framework_TestCase
{

    public function test_it_can_create_csv_files()
    {
        $writer = new DelimitedWriter();

        $writer->insertOne([
            'Column A',
            'Column B',
        ]);


        ob_start();
        $writer->output();
        $csv_data1 = ob_get_clean();


        $writer->insertOne([
            'Column A2',
            'Column B2',
        ]);

        ob_start();
        $writer->output();
        $csv_data2 = ob_get_clean();

        $this->assertEquals("Column A,Column B\n", $csv_data1);
        $this->assertEquals("Column A,Column B\nColumn A2,Column B2\n", $csv_data2);
    }

    public function test_it_can_create_tsv_files()
    {
        $writer = (new DelimitedWriter())->setDelimiter("\t");

        $writer->insertOne([
            'Column A',
            'Column B',
        ]);


        ob_start();
        $writer->output();
        $csv_data1 = ob_get_clean();


        $writer->insertOne([
            'Column A2',
            'Column B2',
        ]);

        ob_start();
        $writer->output();
        $csv_data2 = ob_get_clean();

        $this->assertEquals("Column A\tColumn B\n", $csv_data1);
        $this->assertEquals("Column A\tColumn B\nColumn A2\tColumn B2\n", $csv_data2);
    }
}