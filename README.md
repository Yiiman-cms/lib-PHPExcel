# lib-PHPExcel
Yii2 Excel Functions and class based PHPExcel package
## Install
``composer require yiiman/yii-lib-excel``
## Usage
Create component from ```YiiMan\YiiLibExcel\Excel``` class

### Create Component In Yii2
```php
    'components'=>
    [
    'excel'=>[
        'class'=>\YiiMan\YiiLibExcel\Excel::class
    ]  
];
```

```php
Yii::$app->excel
    ->loadFile(__DIR__.'/list.xlsx')
    ->freezeFirstRow()
    ->setValue(1,2,'Hello excel')
    ->saveAndGetFilePath();
```