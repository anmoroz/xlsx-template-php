<?php
/**
 * Author: Andrey Morozov
 * Date: 22.07.2016
 */

require(__DIR__ . '/../src/Settings.php');
require(__DIR__ . '/../src/LoopData.php');

use XLSXTemplate\Settings;
use XLSXTemplate\LoopData;

class SettingsTest extends PHPUnit_Framework_TestCase
{
    /**
     * @var Settings
     */
    private $settings;

    private $fixture = [
        'providerName' => 'ИП Сумкин Ф.Ф.',
        'clientName' => 'ООО "Рога и копыта"',
        'docNumber' => 5485,
    ];

    public function setUp()
    {
        $this->settings = new Settings($this->fixture);
    }

    public function testLoopData()
    {
        $loopData = new LoopData();
        $this->settings->addLoop('index', $loopData);

        $this->assertInstanceOf('\XLSXTemplate\LoopData', $this->settings->getLoopData('index'));
        $this->assertFalse($this->settings->getLoopData('nonExistentIndex'));
    }

    public function testAddVariable()
    {
        $e = null;
        try {
            $this->settings->addVariable('totalSum', [555]);
        } catch (\InvalidArgumentException $e) {

        }

        $this->assertInstanceOf('\InvalidArgumentException', $e);
    }

    public function testGetValue()
    {
        $this->assertEquals($this->settings->getValue('docNumber'), 5485);
        $this->assertEquals($this->settings->getValue('nonExistentVariable'), '');
        $this->assertEquals($this->settings->getValue('clientName'), 'ООО "Рога и копыта"');

        $this->settings->addVariable('totalSum', function() {
            return 500 + 500;
        });

        $this->assertEquals($this->settings->getValue('totalSum'), 1000);
    }
}