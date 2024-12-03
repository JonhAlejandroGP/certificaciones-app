<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Certificaciones laborales</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f0f0f0;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
        }
        .container {
            background-color: #fff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            width: 350px;
            box-sizing: border-box;
        }
        h2 {
            text-align: center;
            color: #333;
            font-size: 18px;
            margin-bottom: 20px;
        }
        h1 {
            text-align: center;
            color: #333;
            font-size: 20px;
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin: 10px 0 5px;
            color: #555;
            font-size: 14px;
        }
        input[type="text"],
        input[type="file"] {
            width: 100%;
            padding: 10px;
            margin-top: 5px;
            border: 1px solid #ccc;
            border-radius: 5px;
            box-sizing: border-box;
        }
        button {
            width: 100%;
            padding: 10px;
            margin-top: 15px;
            background-color: #004A93;
            border: none;
            border-radius: 5px;
            color: white;
            font-size: 16px;
            cursor: pointer;
            transition: background-color 0.3s ease;
        }
        button:hover {
            background-color: #218838;
        }
        .form-section {
            margin-bottom: 30px;
        }
        .message {
            text-align: center;
            margin-top: 15px;
            color: #333;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>Certificaciones laborales</h1>
    <div class="form-section">
    <h2>Generar por documento</h2>
        <form action="/buscar" method="GET">
            <label for="parametro">Identificación:</label>
            <input type="text" id="parametro" name="parametro">
            <button type="submit">Buscar</button>
        </form>
    
    </div>

    <div class="form-section">
        <h2>Generar con un listado de documentos</h2>
        <form action="/subir" method="POST" enctype="multipart/form-data">
            @csrf
            <label for="archivo">Archivo Excel:</label>
            <input type="file" id="archivo" name="archivo" accept=".xlsx, .xls" required>
            <button type="submit">Subir</button>
        </form>
    </div>
    
    <h2>Generar certificaciones de todos los funcionarios</h2>
    <form action="/buscar" method="GET">
        <input type="hidden" id="parametro" name="parametro" value="0">
        <button type="submit">Generar</button>
        @if(isset($mensaje))
            <p class="message">{{ $mensaje }}</p>
        @endif
    </form>
    
    @if(session('mensaje'))
        <p class="message">{!! session('mensaje') !!}</p>
    @endif
    <?php
    echo date('Y-m-d H:i:s');
    ?>
 </div>
</body>
</html>