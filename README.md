xlsx-template-php
============================
A library for templating of Excel Office documents. Supports MS Excel xlsx. Uses PHPExcel library.

Not only large documents.

An example
------------

#### Input template

![waybill_template](https://raw.github.com/anmoroz/XlsxTemplatePHP/master/demo/waybill_template.jpg)

#### Some code

``` php
$templator = new Templator($templateFile, $outputDir, 'document.xlsx');


$settings = new Settings([
    'providerName' => 'ИП Сумкин Ф.Ф.',
    'clientName' => 'ООО "Рога и копыта"',
    'docNumber' => 5485,
    'docDate' => date('d.m.Y'),
    'totalProductAmount' => 4,
    'totalProductSum' => 2000,
]);


$loopData = new LoopData();
$loopData->setMap([
    'productName',
    'productArticle',
    'productUnit',
    'productAmount',
    'productPrice',
    'productSum',
]);
$loopData->setSource([
   ['Стол письменный', 'СП-234', 'шт.', 1, 1500, 1500],
   ['Чернильница', '75332', 'шт.', 2, 200, 400],
   ['Лампа настольная 12Вт', '6454', 'шт.', 1, 100, 100],
]);
$settings->addLoop(1, $loopData);


$templator->render($settings);
$templator->save();
```

#### Output document

![document](https://raw.github.com/anmoroz/XlsxTemplatePHP/master/demo/document.jpg)