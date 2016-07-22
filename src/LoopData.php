<?php
/**
 * Author: Andrey Morozov
 * Date: 21.07.2016
 */

namespace XLSXTemplate;

class LoopData
{
    /**
     * @var array|object
     */
    private $source;

    /**
     * @var array
     */
    private $map = [];

    /**
     * @param array $map
     */
    public function setMap(array $map)
    {
        $this->map = $map;
    }

    /**
     * @return int
     */
    public function count()
    {
        if (is_array($this->source)) {

            return count($this->source);
        }

        $i = 0;
        foreach ($this->source as $item) {
            $i++;
        }

        return $i;
    }

    /**
     * @return array
     */
    public function getMap()
    {
        return $this->map;
    }

    /**
     * @return array|object
     */
    public function getSource()
    {
        return $this->source;
    }

    /**
     * @param array $source
     */
    public function setSource(array $source)
    {
        if (!is_array($source) && !($source instanceof \Traversable)) {
            throw new \InvalidArgumentException('Source must be array or a class with traversable interface.');
        }
        $this->source = $source;
    }
}