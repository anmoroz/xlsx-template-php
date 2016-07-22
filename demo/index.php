<?php
/**
 * Author: Andrey Morozov
 * Date: 21.07.2016
 */

require(__DIR__ . '/../vendor/autoload.php');
require('Product.php');

use XLSXTemplate\Templator;
use XLSXTemplate\Settings;
use XLSXTemplate\LoopData;

$templateFile = __DIR__.'/templates/waybill_template.xlsx';
$outputDir = __DIR__.'/output/';

$templator = new Templator($templateFile, $outputDir, 'document.xlsx');

$settingsData = [
    'providerName' => 'ИП Сумкин Ф.Ф.',
    'clientName' => 'ООО "Рога и копыта"',
    'docNumber' => 5485,
    'docDate' => date('d.m.Y'),
    'totalProductAmount' => 4,
    'totalProductSum' => function() {
        return 1500 + 400 + 100;
    },
];

$settings = new Settings($settingsData);

$map = [
    'productName',
    'productArticle',
    'productUnit',
    'productAmount',
    'productPrice',
    'productSum',
];
$source = [
    ['Стол письменный', 'СП-234', 'шт.', 1, 1500, 1500],
    ['Чернильница', '75332', 'шт.', 2, 200, 400],
    ['Лампа настольная 12Вт', '6454', 'шт.', 1, 100, 100],
];

$loopData = new LoopData();
$loopData->setMap($map);
$loopData->setSource($source);

$settings->addLoop(1, $loopData);

$templator->render($settings);
$templator->save();


// -------------------------------------------------------
// Source as an array of objects
// -------------------------------------------------------

$templator = new Templator($templateFile, $outputDir, 'document2.xlsx');

$loopData = new LoopData();
$loopData->setMap($map);

$sourceArrayOfObjects = [];
foreach ($source as $sourceItem) {
    $product =  new Product();
    $product->setAttributes($sourceItem);
    $sourceArrayOfObjects[] = $product;
}

$loopData->setSource($sourceArrayOfObjects);

$settings = new Settings($settingsData);
$settings->addLoop(1, $loopData);

$templator->render($settings);
$templator->save();