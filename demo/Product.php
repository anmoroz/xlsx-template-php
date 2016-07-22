<?php
/**
 * Author: Andrey Morozov
 * Date: 22.07.2016
 */

/**
 * Test product class
 *
 * Class Product
 */
class Product
{
    public $productName;
    public $productArticle;
    public $productUnit;
    public $productAmount;
    public $productPrice;
    public $productSum;

    /**
     * Sets the attribute values in a massive way.
     *
     * @param array $values
     */
    public function setAttributes(array $values)
    {
        $attributes = $this->attributes();
        foreach ($values as $key => $value) {
            if (isset($attributes[$key])) {
                $attrName = $attributes[$key];
                $this->$attrName = $value;
            }
        }
    }

    /**
     * Attribute list
     *
     * @return array
     */
    public function attributes()
    {
        return [
            'productName',
            'productArticle',
            'productUnit',
            'productAmount',
            'productPrice',
            'productSum',
        ];
    }
}