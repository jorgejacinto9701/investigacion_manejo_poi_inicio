package com.proyecto.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import com.proyecto.entidad.Alumno;
import com.proyecto.repository.AlumnoRepository;
import com.proyecto.util.Constantes;
import com.proyecto.util.FunctionExcelUtil;
import com.proyecto.util.FunctionUtil;
import com.proyecto.util.UtilExcel;

@Service
public class AlumnoServiceImpl implements AlumnoService {


	@Autowired
	private AlumnoRepository alumnoRepository;

	private static String[] COLUMNAS_Tema = {  "NOMBRE",	"DNI",	"CORREO",  "FECHA NACIMIENTO" };
	private static CellType[][] TIPOS_DATOS_Tema = { {CellType.STRING}, {CellType.STRING, CellType.NUMERIC} , {CellType.STRING}, {CellType.STRING} };

	@Transactional
	@Override
	public Map<String, Object> insertaAlumnoExcel(MultipartFile file) {
		Map<String, Object> mensajes = new HashMap<>();
		Workbook workbook = null;
		InputStream inputStream = null;
		try {
			inputStream = file.getInputStream();
			workbook = new XSSFWorkbook(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
			boolean noEsPrimero = false;

			mensajes = FunctionExcelUtil.validaCabecera(COLUMNAS_Tema, sheet);
			if (mensajes.size() > 0) {
				return mensajes;
			}

			mensajes = FunctionExcelUtil.validaTipoDatosArreglo(TIPOS_DATOS_Tema, sheet);
			if (mensajes.size() > 0) {
				return mensajes;
			}

			List<Alumno> lstAlumnos = new ArrayList<Alumno>(); 
			String celdaNombre = null, celdaDNI = null, celdaCorreo = null, celdaFecha = null;
			for (Row row : sheet) {
				if (noEsPrimero) {
					for (Cell cell : row) {
						switch (cell.getColumnIndex() + 1) {
						case 1:
							celdaNombre = cell.getStringCellValue().trim();
							celdaNombre = FunctionUtil.eliminaEspacios(celdaNombre);
							break;
						case 2:
							if (cell.getCellType() == CellType.NUMERIC) {
								celdaDNI = String.valueOf((int) cell.getNumericCellValue()).trim();
							} else {
								celdaDNI = cell.getStringCellValue().trim();
								celdaDNI = FunctionUtil.eliminaEspacios(celdaDNI);
							}
							break;
						case 3:
							celdaCorreo = cell.getStringCellValue().trim();
							celdaCorreo = FunctionUtil.eliminaEspacios(celdaCorreo);
							break;
						case 4:
							celdaFecha = cell.getStringCellValue().trim();
							celdaFecha = FunctionUtil.eliminaEspacios(celdaFecha);
							System.out.println(celdaFecha);
							break;							
						}
					}
					Alumno objAlumno = new Alumno();
					objAlumno.setNombre(celdaNombre);
					objAlumno.setDni(celdaDNI);
					objAlumno.setCorreo(celdaCorreo);
					objAlumno.setFechaNacimiento(FunctionUtil.getFechaDate(celdaFecha));
					lstAlumnos.add(objAlumno);
				}
				noEsPrimero = true;
			}
			
			int ingresados = 0;
			for (Alumno  objAlumno: lstAlumnos) {
				Alumno objSalida = alumnoRepository.save(objAlumno);
				if(objSalida != null) {
					ingresados++;	
				}
			}
			mensajes.put("mensaje", "Se ha ingresado " + ingresados + " alumno(s).");
			return mensajes;
		} catch (Exception e) {
			e.printStackTrace();
			mensajes.put("mensaje", Constantes.MENSAJE_REG_ERROR);
		} finally {
			try {
				if (inputStream != null)
					inputStream.close();
				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return mensajes;
	}


	@Override
	public List<Alumno> listaAlumno() {
		return alumnoRepository.findAll();
	}
	
	
	private static String[] HEADERs = {"CÓDIGO", "NOMBRE", "DNI", "CORREO", "FECHA NACIMIENTO"};
	private static String SHEET = "Listado";
	private static String TITLE = "Listado de alumnos - Autor: Jorge Jacinto";
	private static int[] HEADER_WITH = { 3000, 10000, 6000, 10000, 8000 };

	public ByteArrayInputStream listaAlumnoExcel(List<Alumno> lstAlumno) {
		try (Workbook excel = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {

			//Se crea la pestaña o hoja
			Sheet hoja = excel.createSheet(SHEET);
			hoja.addMergedRegion(new CellRangeAddress(0, 0, 0, HEADER_WITH.length - 1));

			for (int i = 0; i < HEADER_WITH.length; i++) {
				hoja.setColumnWidth(i, HEADER_WITH[i]);
			}

			CellStyle estiloHeadCentrado = UtilExcel.setEstiloHeadCentrado(excel);
			CellStyle estiloHeadIzquierda = UtilExcel.setEstiloHeadIzquierda(excel);
			CellStyle estiloNormalCentrado = UtilExcel.setEstiloNormalCentrado(excel);
			CellStyle estiloNormalIzquierda = UtilExcel.setEstiloNormalIzquierdo(excel);

			//Fila 0
			Row fila1 = hoja.createRow(0);
			Cell celAuxs = fila1.createCell(0);
			celAuxs.setCellStyle(estiloHeadIzquierda);
			celAuxs.setCellValue(TITLE);

			//Fila 1
			Row fila2 = hoja.createRow(1);
			Cell celAuxs2 = fila2.createCell(0);
			celAuxs2.setCellValue("");

			//Fila 2
			Row fila3 = hoja.createRow(2);
			for (int i = 0; i < HEADERs.length; i++) {
				Cell celda1 = fila3.createCell(i);
				celda1.setCellStyle(estiloHeadCentrado);
				celda1.setCellValue(HEADERs[i]);
			}

			//Fila 3....n
			int rowIdx = 3;
			for (Alumno obj : lstAlumno) {
					Row row = hoja.createRow(rowIdx++);

					Cell cel0 = row.createCell(0);
					cel0.setCellValue(obj.getIdAlumno());
					cel0.setCellStyle(estiloNormalCentrado);

					Cell cel1 = row.createCell(1);
					cel1.setCellValue(obj.getNombre());
					cel1.setCellStyle(estiloNormalIzquierda);
	
					Cell cel2 = row.createCell(2);
					cel2.setCellValue(obj.getDni());
					cel2.setCellStyle(estiloNormalCentrado);
					
					Cell cel3 = row.createCell(3);
					cel3.setCellValue(obj.getCorreo());
					cel3.setCellStyle(estiloNormalIzquierda);
					
					Cell cel4 = row.createCell(4);
					cel4.setCellValue(FunctionUtil.getFechaString(obj.getFechaNacimiento()));
					cel4.setCellStyle(estiloNormalCentrado);
			}
			excel.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		} catch (IOException e) {
			throw new RuntimeException("Error en importar data en el Excel: " + e.getMessage());
		}
	}

	
}
