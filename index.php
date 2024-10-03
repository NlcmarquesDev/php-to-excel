<?php

declare(strict_types=1);
require __DIR__ . '/vendor/autoload.php';

use src\Classes\Excel;

//Aqui foi so um teste com  a o hardeCode inserido com array, mas pode ser atraves do post request que vem atraves do formulário

$newFile = new Excel('test1', 'test-excel', 'newfile excell', 'apenas para test se trabalha');

$newFile->getSpreadsheet('Aluno', [
  'A' => 'Nome',
  'B' => 'Idade',
  'C' => 'Email',
]);

$newFile->getProperties([
  [
    'A' => 'Ines',
    'B' => '19',
    'C' => 'ines@.com',

  ],
  [
    'A' => 'Nuno',
    'B' => '20',
    'C' => 'nunoc@.com',
  ]
]);

$newFile->saveAs('test');


?>
<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Create a document to excell</title>
</head>

<body>
  <h1>Create a document to excel</h1>
  <form action="" method="post">
    <div>
      <label for="">Name: </label>
      <input type="text" name="name" id="">
    </div>
    <div>
      <label for="">Idade: </label>
      <input type="text" name="age" id="">
    </div>
    <div>
      <label for="">Email: </label>
      <input type="text" name="email" id="">
    </div>
    <div>
      <input type="submit" value="Create">
    </div>
    </div>

  </form>
</body>

</html>