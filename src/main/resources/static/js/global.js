var regexNotaValida = /^((\d\.\d{0,1})|((1)\d)|(\d)|(20)|((1)\d\.\d{0,1}))$/;
var regexNotaEnvio = /^((\d\.\d{1})|((1)\d)|(\d)|(20)|((1)\d\.\d{1}))$/;
var MSG_ELIMINAR_MIGRACION = "&#191;Desea eliminar todos los datos migrados&#63;";
var MSG_ELIMINAR_NOTA = "&#191;Desea eliminar todas las notas&#63;";
var MSG_ELIMINAR = "&#191;Desea activar o desactivar el registro&#63;";
var MSG_ELIMINAR_FISICA = "&#191;Desea eliminar el registro&#63;";
var MSG_REGISTRAR_NOTA = "&#191;Desea registrar notas&#63;";
var MSG_RECTIFICAR_REGISTRA_DOCUMENTO = "&#191;Desea registrar el ";
var MSG_RECTIFICAR_ACTUALIZA_DOCUMENTO = "&#191;Desea actualizar el ";
var MSG_REGISTRAR_NOTAS_VACIAS = "Todas las notas est&aacute;n vacias";
var MSG_REGISTRAR_NOTA_FORMATO = "La nota debe ser del 1 al 20 entero o con un decimal";
var MSG_REGISTRAR_CONVALIDACION_TODOS_CURSOS = "Debe seleccionarse todos los cursos";
var MSG_REGISTRAR_ASISTENCIA = "&#191;Desea registrar asistencia&#63;";
var MSG_ERROR = "Se gener&oacute; un error, consulte al administrador";
var MSG_RANGO_FECHA = "Fecha de fin debe ser mayor o igual a la fecha de inicio";
var MSG_RANGO_FECHA_NOTA = "Fecha de fin de registro de notas debe ser mayor o igual a la fecha de inicio de registro de notas";
var MSG_RANGO_FECHA_DISPONIBILIDAD = "Fecha de fin de registro de disponibilidad debe ser mayor o igual a la fecha de inicio de registro de disponibilidad";
var MSG_RANGO_FECHA_ANHO = "El rango de fecha no debe ser mayor a un a&ntilde;o";
var MSG_SESIONES_POGRAMADAS = "Ud. no tiene sesiones programadas para hoy.";
var MSG_TURNO_POGRAMADOS = "Ud. no tiene turno programados";
var MSG_SESIONES_POGRAMADAS_EN_PERIODO = "Ud. no tiene sesiones programadas en el periodo.";
var MSG_NOTAS_EN_PERIODO = "No tiene notas en el periodo.";
var MSG_DISPONIBILIDAD_EN_PERIODO  = "No existe disponibilidad ingresada";
var MSG_HORARIOS_POGRAMADAS_EN_PERIODO = "No existe horarios programados para esa consulta.";
var MSG_HORARIO_VALIDA_SELECCIONES = "Seleccione las clases a actualizar";
var MSG_HORARIO_VALIDA_CONFIRMACION = "&#191;Desea actualizar las clases seleccionadas&#63;";
var MSG_VALIDA_BUSCAR = "Pulse Buscar antes de Registrar";
var MSG_VALIDA_MODALIDAD = "Seleccione la Modalidad";

var MSG_SEGURIDAD_ROL = "Seleccione el Rol";
var MSG_SEGURIDAD_OPCION_AGREGAR = "Seleccione una o varias opciones a agregar";
var MSG_SEGURIDAD_OPCION_ELIMINAR = "Seleccione una o varias opciones a eliminar";

var MSG_SEGURIDAD_USUARIO = "Seleccione el Usuario";
var MSG_SEGURIDAD_ROL_AGREGAR = "Seleccione una o varios roles a agregar";
var MSG_SEGURIDAD_ROL_ELIMINAR = "Seleccione una o varios roles a eliminar";

var MSG_VALIDA_FILE_XLXS = "El archivo debe tener extensi&oacute;n .xlsx";
var MSG_VALIDA_FILE_IMAGEN = "El archivo debe tener extensi&oacute;n .jpg, .png o .jpeg";
var MSG_VALIDA_FILE_TAMANO_UN_MB = "El archivo debe tener 1MB";
var MSG_VALIDA_FILE_NOMBRE_TAMANO = "El nombre del archivo debe tener hasta 100 caracteres";
var MSG_VALIDA_FILE_TAMANO_MEDIO_MEGA = "El archivo debe tener 500KB";
var MSG_VALIDA_FILE_TAMANO_UN_MEGA = "El archivo debe tener 1MB";

var MSG_VALIDA_FILE_EXISTENCIA = "Seleccione el archivo";
var MSG_VALIDA_FILE_ERROR_NOMBRE = "El archivo debe tener el nombre : ";

var MSG_MATRICULA_CURSO = "Seleccione el curso";
var MSG_MATRICULA_HORARIO = "Seleccione el horario";
var MSG_MATRICULA_ELIMINAR_TODOS_HORARIO = "Desea eliminar todos los horarios seleccionados";
var MSG_MATRICULA_ELIMINAR_UN_HORARIO = "Desea eliminar el horario seleccionado";
var MSG_MATRICULA_SIN_HORARIO = "No existe horarios";

var TIPO_PAGO_PENDIENTE = 42;
var TIPO_PAGO_PAGADO = 43;

var NOTA_NO_RINDIO = 77;
var NOTA_NO_PROCESADA = 88;

var NOTA_NO_REGISTRADA = 36;
var NOTA_SI_REGISTRADA = 37;
var NOTA_PRE_INGRESADA = 39;

var ESTADO_DICTADO_NO_DICTADO = 4;
var ESTADO_DICTADO_SI_DICTADO = 3;

var ASISTENCIA_ALUMNO_SI_ASISTIO = 10;
var ASISTENCIA_ALUMNO_TARDANZA	 = 11;
var ASISTENCIA_ALUMNO_NO_ASISTIO = 12;

var MSG_VALIDA_NOTA_INGRESADA = "Las notas deben estar ingresadas para la rectificaci&oacute;n";
var MSG_FILE_UPLOAD = "El archivo se est&aacute; procesando...";

var INACTIVO = 0;
var ACTIVO = 1;

var NO_VIGENTE = 0;
var VIGENTE = 1;

var MSG_OK = "OK";
var MSG_ERROR = "ERROR";

var CURSO_OBLIGATORIO = "Si";
var CURSO_OPCIONAL = "No";

var TIPO_NO_CLASIFICADOS = 0;
var TIPO_SI_CLASIFICADOS = 1;

function getTipoMensaje(val){
	if (val == "Clasificado") return 1;
	else return 0;
}

function getTipoPuntuacion(val){
	if (val == 1) return "Positivo";
	else if (val == 2) return "Neutro";
	else if (val == 3) return "Negativo";
}

function getTipoMensaje(val){
	if (val == 0) return "No Clasificado";
	else if (val == 1) return "NA";
	else if (val == 2) return "Valido";
}

var IDIOMA = {
	processing:"procesando...",
    lengthMenu: "_MENU_ Registros por p&aacute;gina",
    zeroRecords: "No existen registros",
    info: "P&aacute;gina _PAGE_ de _PAGES_",
    infoEmpty: "Sin registros",
    infoFiltered: "(Filtro de _MAX_ registros)",
    search: "Buscar:",
    paginate: {
        "first":      "Primero",
        "last":       "Last",
        "next":       "Siguiente",
        "previous":   "Anterior"
    }
};


var tooltipCloseHandler;
var closeTooltip = function (element) {
    if (element) {
        $(element).popover("destroy");
        currentPopoverElement = null;
    } else {
        if ($(".popover .modal-select").hasClass("open")) {
            $(".popover .dropdown-toggle").trigger("click");
        }
        $('.popover').remove();
    }
    $(document).off("click", "body", tooltipCloseHandler);
};

var openTooltip = function (element, title, html) {
    closeTooltip(element);
    currentPopoverElement = element;
    $(element).popover({
        html: true,
        placement: 'auto top',
        trigger: 'manual',
        title: title,
        content: html,
        container: 'body'
    });
    $(element).popover("show");

    tooltipCloseHandler = $(document).on('mousedown', "body", function (e) {
        if (!$(e.target).is(".popover") && $(".popover").has(e.target).length === 0) {
            closeTooltip();
        }
    });
};

var CALENDAR_HORARIO = {
		 lang: 'es',	
		 header: {left: '', center: '',	right: ''},
		 columnFormat: 'dddd M/D',
		 minTime: '07:00:00',
		 maxTime: '24:00:00',
		 contentHeight: 'auto',
		 allDaySlot:false, 
		 defaultView: 'agendaWeek',
		 defaultDate: new Date(), 
		 navLinks: false, 
		 selectable: false, 
		 selectHelper: true,
		 select: function(start, end, allDay) {},
		 editable: false, 
		 eventLimit: true, 
		 axisFormat: 'h(:mm)a',
		 events: [],
		 eventClick: function (event, jsEvent, view) {
             closeTooltip();
             var current = new Date();
             openTooltip(jsEvent.target, '<strong>' + event.curso + '</strong>', Mustache.render($('#idPlantillaHorario').html(), {
                 date: event.start.format('DD/MM/YYYY'),
                 startTime: event.start.format('hh:mm a'),
                 endTime: event.end.format('hh:mm a'),
                 docente: event.docente,
                 asignatura: event.curso,
                 seccion: event.seccion,
                 tipo: "Curso:",
                 aula: event.aula || 'Ninguna',
                 editable: event.source.sourceType === 'main'
             }));
         },
}


var CALENDAR_HORARIO_MATRICULA = {
		 lang: 'es',	
		 header: {left: '', center: '',	right: ''},
		 columnFormat: 'dddd',
		 minTime: '07:00:00',
		 maxTime: '24:00:00',
		 contentHeight: 'auto',
		 allDaySlot:false, 
		 defaultView: 'agendaWeek',
		 defaultDate: new Date(), 
		 navLinks: false, 
		 selectable: false, 
		 selectHelper: true,
		 select: function(start, end, allDay) {},
		 editable: false, 
		 eventLimit: true, 
		 axisFormat: 'h(:mm)a',
		 events: [],
		 eventClick: function (event, jsEvent, view) {
            closeTooltip();
            var current = new Date();
            openTooltip(jsEvent.target, '<strong>' + event.curso + '</strong>', Mustache.render($('#idPlantillaHorario').html(), {
                date: event.start.format('DD/MM/YYYY'),
                startTime: event.start.format('hh:mm a'),
                endTime: event.end.format('hh:mm a'),
                docente: event.docente,
                asignatura: event.curso,
                seccion: event.seccion,
                tipo: "Curso:",
                aula: event.aula || 'Ninguna',
                editable: event.source.sourceType === 'main'
            }));
        },
}

function mostrarMensajeUpload(msg){
	$('#id_my_modal_upload').remove();
	$("body").append('<div class="modal fade" id="id_my_modal_upload" role="dialog"><div class="modal-dialog"><div class="modal-content">' + 
					 '<div class="modal-header"><h4 class="modal-title">Subiendo Archivo</h4></div>'+
					 '<div class="modal-body"><p>'+msg+'</p></div>'+
					 '</div></div></div>');
	$('#id_my_modal_upload').modal("show");
}

function mostrarMensaje(msg, ejecutarAccion){
	$('#id_my_modal').remove();
	
	$("body").append('<div class="modal fade" id="id_my_modal" role="dialog"><div class="modal-dialog"><div class="modal-content">' + 
					 '<div class="modal-header"><button type="button" class="close" data-dismiss="modal">&times;</button><h4 class="modal-title">Mensaje</h4></div>'+
					 '<div class="modal-body"><p>'+msg+'</p></div>'+
					 '<div class="modal-footer"><button type="button" id="id_btn_dialog_aceptar" class="btn btn-primary btn-sm">Aceptar</button></div>'+
					 '</div></div></div>');
	
	$("#id_btn_dialog_aceptar").click(function(){
		$('#id_my_modal').modal("hide");
		if(ejecutarAccion != undefined){
			window.location.href = ejecutarAccion;
		}
	});
	
	$('#id_my_modal').modal("show");
}

function mostrarMensajeConfirmacion(msg, accionAceptar, accionCancelar, data){
	$('#id_my_modal_confirmacion').remove();
	
	$("body").append('<div class="modal fade" id="id_my_modal_confirmacion" role="dialog"><div class="modal-dialog"><div class="modal-content">' + 
					 '<div class="modal-header"><button type="button" class="close" data-dismiss="modal">&times;</button><h4 class="modal-title">Mensaje</h4></div>'+
					 '<div class="modal-body"><p>'+msg+'</p></div>'+
					 '<div class="modal-footer">'+
					 '<button type="button" id="id_btn_dialog_si" class="btn btn-primary btn-sm">Si</button>'+
					 '<button type="button" id="id_btn_dialog_no" class="btn btn-default">No</button>'+
					 '</div>'+
					 '</div></div></div>');
	
	$("#id_btn_dialog_si").click(function(){
		$('#id_my_modal_confirmacion').modal("hide");
		if(accionAceptar != null){
			accionAceptar(data);
		}
	});
	$("#id_btn_dialog_no").click(function(){
		$('#id_my_modal_confirmacion').modal("hide");
		if(accionCancelar != null){
			accionCancelar(data);
		}
	});
	
	$('#id_my_modal_confirmacion').modal("show");
}


function getEstado(val){
	if (val == 1) return 'Activo';
	else if (val == 0) return 'Inactivo';
}

function getVigencia(val){
	if (val == 1) return 'Vigente';
	else if (val == 0) return 'No vigente';
}

function getEstadoMatricula(val){
	if (val == 1) return 'En matr&iacute;cula';
	else if (val == 0) return 'No habilitado';
}

function getSexo(val){
	if (val == 1) return 'Masculino';
	else if (val == 2) return 'Femenino';
}

function getTextoBotonEstado(val){
	if (val == 1) return 'Desactivar';
	else if (val == 0) return 'Activar';
}

function getTextoBotonVigente(val){
	if (val == 1) return 'Hacer No Vigente';
	else if (val == 0) return 'Hacer Vigente';
}

function getTextoMatricula(val){
	if (val == 1) return 'Matriculado';
	else if (val == 0) return 'Sin Matricula';
}

function getTextoMatricula(val){
	if (val == 1) return 'Matriculado';
	else if (val == 0) return 'Sin Matricula';
}

function getCambioEstado(val){
	if (val == 1) return 0;
	else if (val == 0) return 1;
}

function getNotaRendida(val){
	if (val == NOTA_NO_REGISTRADA || val == NOTA_NO_RINDIO) return "NR";
	else return val;
}

function getPesoPorcentaje(val){
	return (val *100) + "%";  
}

function getTardanza(val){
	if (val == -1) return "-";
	else return val;
}

function getFechaVacio(val){
	if (val == "") return "-";
	else return val;
}

function getFormatoFormula(val){
	return val.replaceAll("+"," + ");
}

function getCursoEnCurriculaEstado(val){
	if (val == 0) return "Pendiente";
	else return "Cumplido";
}

function valida_rango(xinput, pnotamax) {
    var xkey = event.keyCode;
    var valor;
    if (((xkey != 46) && (xkey < 48)) || ((xkey > 57))) {
        if (event.preventDefault) {
            event.preventDefault();
        }
        else {
            event.returnValue = false;
        }
    }
    valor = xinput.value + String.fromCharCode(xkey);
    if (valor > pnotamax) {
        if (event.preventDefault) {
              event.preventDefault();
        }else {
              event.returnValue = false;
        }
    }

    if (!regexNotaValida.test(valor)){
      if (event.preventDefault) {
            event.preventDefault();
      }else {
            event.returnValue = false;  
      }
    }
}

function getFechaActual(){
    var val = new Date();
    return  dateToString(val); 
}

function getFechaInicioAnhoActual(){
    var val = new Date();
    var inicioAnho = new Date();
    inicioAnho.setFullYear(val.getFullYear(), 0, 1);
    return  dateToString(inicioAnho); 
}

function dateToString(date){
    var formattedDate= new Date(date);
    var d = formattedDate.getDate();
    var m =  formattedDate.getMonth();
    m += 1;
    var y = formattedDate.getFullYear();
    return y + "-" + lpads(m,2) + "-" + lpads(d,2);
}

function lpads(n, width, z) {
    z = z || '0';
    n = n + '';
    return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
}

function valFechaInicioMayFechaFin(idIni, idFin){
    var fIni = $.trim($(idIni).val());
    var fFin = $.trim($(idFin).val());
    
    var dIni = new Date(fIni);
    var dFin = new Date(fFin);
    
    if (dIni > dFin){
        return true;
    }else{
        return false;
    }
}

function valFechaFinMayorAnhoFecInicio(idIni, idFin){
    var fIni = $.trim($(idIni).val());
    var fFin = $.trim($(idFin).val());
    
    var dIni = new Date(fIni);
    var dFin = new Date(fFin);
    
    var proximoAnho = new Date();
    proximoAnho.setFullYear(dIni.getFullYear() + 1, dIni.getMonth(), dIni.getDate());

    if (dFin > proximoAnho) {
         return true;
    }else{
        return false;
    }
}
/*
if(history.forward(1)) {
	location.replace(history.forward(1));
}

$(function(){
    $(document).bind("contextmenu",function(e){
        return false;
    });
});
*/

function getStringDia(pos) {
	if (pos == 1)return "Domingo";
	else if (pos == 2)return "Lunes";
	else if (pos == 3)return "Martes";
	else if (pos == 4)return "Mi&eacute;rcoles";
	else if (pos == 5)return "Jueves";
	else if (pos == 6)return "Viernes";
	else return "S&aacute;bado";
}

function valida_numero() {
    var xkey = event.keyCode;
    if (((xkey < 48)) || ((xkey > 57))) {
        if (event.preventDefault) {
            event.preventDefault();
        }
        else {
            event.returnValue = false;
        }
    }
}
