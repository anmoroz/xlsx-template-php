<?php
/**
 * Author: Andrey Morozov
 * Date: 22.07.2016
 */

require(__DIR__ . '/../src/LoopData.php');

use XLSXTemplate\LoopData;

class LoopDataTest extends PHPUnit_Framework_TestCase
{
    /**
     * @var LoopData
     */
    private $loopData;

    private $fixtureSource = [
        ['Стол письменный', 'СП-234', 'шт.', 1, 1500, 1500],
        ['Чернильница', '75332', 'шт.', 2, 200, 400],
        ['Лампа настольная 12Вт', '6454', 'шт.', 1, 100, 100],
    ];

    public function setUp()
    {
        $map = [
            'productName',
            'productArticle',
            'productUnit',
            'productAmount',
            'productPrice',
            'productSum',
        ];

        $this->loopData = new LoopData();
        $this->loopData->setMap($map);
        $this->loopData->setSource($this->fixtureSource);
    }

    public function testCount()
    {
        $this->assertEquals($this->loopData->count(), 3);
    }

    public function testGetMap()
    {
        $map = $this->loopData->getMap();
        $this->assertEquals(array_search('productArticle', $map), 1);
        $this->assertEquals(count($map), 6);
    }

    public function testGetSource()
    {
        $this->assertEquals($this->fixtureSource, $this->loopData->getSource());
    }
}