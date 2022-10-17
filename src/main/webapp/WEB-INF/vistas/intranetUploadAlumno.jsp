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
<div class="container" style="margin-top: 4%"><h4>Ingreso de Alumno</h4></div>
<div class="container" style="margin-top: 1%">
	<form  id="id_form" method="post" enctype="multipart/form-data">
			<div class="row"  style="margin-top: 2%">
		            <label class="col-lg-3 control-label" for="id_formato">Archivo Excel: &nbsp;&nbsp; <img alt="Excel" src="img/excel.png" width="22px" height="22px"> </label>
		            <div class="col-lg-5">
						<input class="form-control" id="id_formato"name="formato" placeholder="Ingrese el criterio" type="file"/>
		            </div> 
		       </div>
		     <div class="row"  style="margin-top: 2%">
		            <div class="col-lg-8" align="center">
						<button  type="button"  id="id_btn_subir"  class="btn btn-primary btn-sm" style="width: 150px;">
							<i style="font-size: 15px;" class="fa fa-upload"></i> Subir Excel
						</button>
		            </div>
		  	</div>
	</form>
</div>

<script type="text/javascript">
$("#id_btn_subir").click(function(){
    var fileUrl = $("#id_formato").val();
    if(fileUrl == ''){
        mostrarMensaje(MSG_VALIDA_FILE_EXISTENCIA);
        return false;		
    }
    var extension = fileUrl.split('.').pop().toLowerCase();
    if(extension.toLowerCase() != 'xlsx'){
      mostrarMensaje(MSG_VALIDA_FILE_XLXS);
      return false;		
    }
    var formData = new FormData();
    var file = $('#id_formato')[0].files[0];
    formData.append("file", file);

    $('#id_btn_subir').attr('disabled','disabled');
    mostrarMensajeUpload(MSG_FILE_UPLOAD);
    
    $.ajax({
          type: "POST",
          url: "subirPlantillaTema", 
          data: formData,
          enctype : 'multipart/form-data',
          contentType : false,
          processData : false,
          cache:false,
          success: function(data){
       		  $('#id_my_modal_upload').modal("hide");
       		  $('#id_btn_subir').removeAttr('disabled');
       		  mostrarMensaje(data.mensaje);
       	  },
          error: function(){
        	  $('#id_btn_subir').removeAttr('disabled');
        	  $('#id_my_modal_upload').modal("hide");
        	  mostrarMensaje(MSG_ERROR);
          }
    });
});

</script>   		
</body>
</html>