<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="shortcut icon" href="/IMAGENES/logoElis.png" type="image/x-icon">
  <title>ELIS SYSTEMS HOME</title>
  <style>
    body {
    font-family: Arial, sans-serif;
    background-color: #f4f4f4;
    margin: 0;
    padding: 0;
    display: flex;
    justify-content: center;
    align-items: center;
    background-image: url(https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRaxsGBNxX80asg9URQZhSph3dCmdtGD01wpeEN_cqjSg&s);
    }

    #sidebar {
    margin-top: 70px;
    margin-left: 20px;
      width: 350px;
      height: auto;
      background-color: #333;
      color: #fff;
      position: fixed;
      left: 0;
      top: 0;
      overflow-x: hidden;
      padding-top: 20px;
      background-color:#01a5abb6;
      border-radius: 15px;
    }

    .main-button {
    margin-left: 8px;
    margin-right: 10px;
    margin-bottom: 20px;
    background-color: #e4000d;
      display: block;
      width: 90%;
      padding: 10px;
      border: none;
      text-align: center;
      font-size: 16px;
      cursor: pointer;
      position: relative;
      border-radius: 20px;
    }

    .main-button:hover + .sub-buttons,
    .sub-buttons:hover {
      display: block;
    }

    .sub-button button {
      display: block;
      width: 100%;
      padding: 8px;
      margin: 0;
      border: none;
      text-align: left;
      font-size: 14px;
      cursor: pointer;
      background-color: #01a5abb6;
      border-radius: 15px;
      color: white;
      margin-bottom: 10px;
      margin-top: 8px;
    }

    .sub-button:hover {
      border-radius: 15px;
      background-color: #999;
    }
  </style>
</head>
<body>

<div id="sidebar">
<div class="main-button">
    Actas de Entrega
  <div class="sub-button">
   <button onclick="location.href='./ACTAS COMPUTADORES/ActaEntregaComputadores.php'">Acta de entrega Computadores</button> 
  </div>
  <div class="sub-button">
   <button onclick="location.href='./ACTAS CELULARES/ActaEntregaCelulares.php'">Acta de entrega Celulares</button> 
  </div>
</div>

<div class="main-button">
    Actas Recibido
  <div class="sub-button">
   <button onclick="location.href='./ACTAS COMPUTADORES/ActaRecibidoComputadores.php'">Acta de recibido Computadores</button> 
  </div>
  <div class="sub-button">
   <button onclick="location.href='./ACTAS CELULARES/ActaRecibidoCelulares.php'">Acta de recibido de Celulares</button> 
  </div>
</div>

<div class="main-button">
    Filtrar y Reportar
  <div class="sub-button">
   <button onclick="location.href='./PRUEBA FILTRO CON REPORTES/filtro.php'">Filtro y Reportes</button> 
  </div>
</div>

</div>

</div>

<div id="content">
  <h1>Bienvenido A ELIS SYSTEMS HOME</h1>
  <p>Bienvenido al sistema automatizado de papeleo de elis</p>
</div>

</body>
</html>