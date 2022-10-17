package com.proyecto.service;

import java.io.ByteArrayInputStream;
import java.util.List;
import java.util.Map;

import org.springframework.web.multipart.MultipartFile;

import com.proyecto.entidad.Alumno;

public interface AlumnoService {

	public abstract Map<String, Object> insertaAlumnoExcel(MultipartFile file);
	public abstract List<Alumno> listaAlumno();
	public abstract ByteArrayInputStream listaAlumnoExcel(List<Alumno> lstAlumno);
}
