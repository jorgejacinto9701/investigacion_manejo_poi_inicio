package com.proyecto.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.proyecto.entidad.Alumno;

public interface AlumnoRepository extends JpaRepository<Alumno, Integer>{
	
	
}
