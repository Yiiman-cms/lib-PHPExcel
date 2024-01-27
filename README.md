# lib-PHPExcel
Yii2 Excel Functions and class based PHPExcel package
## Install
``composer require yiiman/yii-lib-excel``
## Usage
Create component from ```YiiMan\\Excel\Excel``` class

### Create Component In Yii2
```php
    'components'=>
    [
    'excel'=>[
        'class'=>YiiMan\\Excel\Excel::class
    ]  
];
```

```php
Yii::$app->excel
    ->loadFile(__DIR__.'/list.xlsx')
    ->freezeFirstRow()
    ->setValue(1,2,'Hello excel')
    ->write_and_get_file_path();
```