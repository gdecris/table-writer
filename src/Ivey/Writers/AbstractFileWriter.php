<?php


namespace Ivey\Writers;


use Illuminate\Support\Collection;

abstract class AbstractFileWriter
{

    protected $data;

    public function __construct()
    {
        $this->data = new Collection();
    }

    /**
     * @param $row
     * @return $this
     */
    public function insertOne($row)
    {
        $this->data->push($row);

        return $this;
    }

    /**
     * @param $rows
     * @return $this
     */
    public function insertAll($rows)
    {
        $this->data = $this->data->merge($rows);

        return $this;
    }

    /**
     * @param null $filename
     */
    abstract public function output($filename = null);
}