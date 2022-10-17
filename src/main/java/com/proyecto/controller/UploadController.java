package com.proyecto.controller;

import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import com.proyecto.service.AlumnoService;

import lombok.extern.apachecommons.CommonsLog;

@Controller
@CommonsLog
public class UploadController {
	
	@Autowired
	private AlumnoService alumnoService;
	
	@RequestMapping("/verUplaodAlumno")
	public String verUplaodAlumno() {	return "intranetUploadAlumno";	}
	
	@PostMapping(value = "/subirPlantillaAlumno", consumes = "multipart/form-data")
	@ResponseBody
	public Map<String, Object> subirPlantillaAlumno(@RequestParam("file") MultipartFile file) {
		log.info(">>> subirPlantillaAlumno ");
		Map<String, Object> salida =alumnoService.insertaAlumnoExcel(file);
		return salida;
	}
	
	
	
}
