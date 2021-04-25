<!DOCTYPE html>
<html lang="en" dir="ltr">
  <head>
    <meta charset="utf-8">
    <title></title>
  </head>
  <body onload="iniciar();">
    <p id="msg"> Descargando archivos espere un momento por favor.... </p>

    <script src="https://code.jquery.com/jquery-3.6.0.js" integrity="sha256-H+K7U5CnXl1h5ywQfKtSj8PCmoN9aaq30gDh27Xc0jk=" crossorigin="anonymous"></script>
    <script type="text/javascript">
    function iniciar(){

      $.ajax({
        method: "POST",
        url: "descarga.php",
      }).done(function( msg ) {
         $("#msg").html(msg);
      });

    }
    </script>

  </body>
</html>
