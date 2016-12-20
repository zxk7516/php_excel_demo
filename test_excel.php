<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Document</title>
    <style>
        table{ border-collapse: collapse}
        td,th,table{border: 1px solid darkgray}
    </style>
</head>
<body>
<?php
require_once './lib/phpexcel/Excel2HtmlRender.php';
//$render = new Excel2HtmlRender('./Engineer.xls');
$render = new Excel2HtmlRender('./test.xlsx');
print $render->render(Excel2HtmlRender::optimize_options(array(
    'renderer'=>'simple',//只能用simple
)));
?>
</body>
</html>


