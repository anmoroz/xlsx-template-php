<?php
/**
 * Author: Andrey Morozov
 * Date: 21.07.2016
 */

namespace XLSXTemplate;

class Settings
{
    /**
     * @var array
     */
    private $loops = [];

    /**
     * @var array
     */
    private $mapVariable = [];

    public function __construct(array $mapVariable = [])
    {
        if (!empty($mapVariable)) {
            foreach ($mapVariable as $key => $value) {
                $this->addVariable($key, $value);
            }
        }
    }

    /**
     * @param string $key
     * @param LoopData $loopData
     */
    public function addLoop($key, LoopData $loopData)
    {
        $this->loops[$key] = $loopData;
    }

    /**
     * @param string $key
     * @return bool|LoopData
     */
    public function getLoopData($key)
    {
        if (isset($this->loops[$key])) {

            return $this->loops[$key];
        }

        return false;
    }

    /**
     * @param string $key
     * @return string
     */
    public function getValue($key)
    {
        if (isset($this->mapVariable[$key])) {
            if (is_scalar($this->mapVariable[$key])) {
                return $this->mapVariable[$key];
            } else {
                return $this->mapVariable[$key]();
            }
        }
        return '';
    }

    /**
     * @param string $key
     * @param $value
     */
    public function addVariable($key, $value)
    {
        if (!is_scalar($value) && !is_callable($value)) {
            throw new \InvalidArgumentException('Value must be scalar or callable.');
        }

        $this->mapVariable[$key] = $value;
    }
}