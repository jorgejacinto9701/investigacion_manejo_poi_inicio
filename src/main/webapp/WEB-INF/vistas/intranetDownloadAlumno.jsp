<!DOCTYPE html>
<html lang="esS" >
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta http-equiv="Expires" content="-1" />
<meta http-equiv="Cache-Control" content="private" />
<meta http-equiv="Cache-Control" content="no-store" />
<meta http-equiv="Pragma" content="no-cache" />

<script type="text/javascript" src="js/jquery.min.1.9.1.js"></script>
<script type="text/javascript" src="js/jquery-ui.min.js"></script>
<script type="text/javascript" src="js/jquery.fileDownload.js"></script>

<script type="text/javascript" src="js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="js/dataTables.bootstrap.min.js"></script>
<script type="text/javascript" src="js/bootstrap.min.js"></script>
<script type="text/javascript" src="js/bootstrapValidator.js"></script>
<script type="text/javascript" src="js/global.js"></script>

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
<link rel="stylesheet" href="css/jquery-ui.css"  />

<link rel="stylesheet" href="css/bootstrap.css"/>
<link rel="stylesheet" href="css/dataTables.bootstrap.min.css"/>
<link rel="stylesheet" href="css/bootstrapValidator.css"/>

<title>Sistema</title>
</head>
<body>
<div class="container" style="margin-top: 4%"><h4>Descarga de Alumno</h4></div>
<div class="container" style="margin-top: 1%">
		 <div class="col-md-12" >
		       <form  action="listaAlumnosDescarga" class="fileDownloadForm"  method="post">
				<div class="row">
						<div class="form-group">
							<button  type="submit" style="width: 220px" class='btn btn-primary btn-sm'>
									<i style="font-size: 15px;" class="fa fa-download"></i> Descarga Excel</button>
							<br>
						</div>
				</div>
		 		</form>
		  </div>
</div>
<script type="text/javascript">

$(document).on("submit", "form.fileDownloadForm", function (e) {
    $.fileDownload($(this).prop('action'), {
        preparingMessageHtml: "Estamos preparando su informe, por favor espere",
        failMessageHtml: "Se produjo un problema al generar su informe. Vuelva a intentarlo.",
        httpMethod: "POST",
        data: $(this).serialize()
    });
    e.preventDefault(); 	
});

</script>


</body>
</html> 