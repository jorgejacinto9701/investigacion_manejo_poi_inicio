package com.proyecto.controller;

import java.io.ByteArrayInputStream;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import com.proyecto.entidad.Alumno;
import com.proyecto.service.AlumnoService;
import com.proyecto.util.Constantes;
import com.proyecto.util.FunctionUtil;

import lombok.extern.apachecommons.CommonsLog;
@CommonsLog
@Controller
public class DownloadController {

	@Autowired
	private AlumnoService alumnoService;
	
	@RequestMapping("/verDownloadAlumno")
	public String verDownloadAlumno() {	return "intranetDownloadAlumno";	}
	
	
	
	@PostMapping("/listaAlumnosDescarga")
	public ResponseEntity<Resource> listaAlumnosDescarga() {
		log.info(">>> listaAlumnosDescarga ");
		
		ByteArrayInputStream bytes = null;
		InputStreamResource body = null;
		String filename  = null;
		try {
			List<Alumno> lista = alumnoService.listaAlumno();
			filename = "Planilla_Reporte_NPS__"+ FunctionUtil.getDateNowInString()+".xlsx";
			bytes = alumnoService.listaAlumnoExcel(lista);
			body = new InputStreamResource(bytes);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return ResponseEntity.ok()
				 	.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)		        		
				 	.header(HttpHeaders.SET_COOKIE, "fileDownload=true; path=/")
				 	.header(HttpHeaders.CACHE_CONTROL, "no-cache, no-store, must-revalidate")
				 	.header(HttpHeaders.CONTENT_LENGTH, String.valueOf(bytes.available()))
			        .contentType(MediaType.parseMediaType(Constantes.TYPE_EXCEL))
			        .body(body);
		
	 }
	
	
}
