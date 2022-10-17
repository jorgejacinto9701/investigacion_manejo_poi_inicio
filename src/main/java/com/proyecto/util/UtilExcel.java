package com.proyecto.util;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class UtilExcel {
	
	public static final String GUION = "-";
	

	public static String validaVacio(String s) {
		if (s == null || s.trim().equals("")) {
			return "-";
		} else {
			return s;
		}
	}
	
	public static String validaMenosUno(int s) {
		if (s == -1) {
			return "-";
		} else {
			return String.valueOf(s);
		}
	}
	
	public static String validaEstado(int s) {
		if (s == 0) {
			return "Inactivo";
		} else {
			return "Activo";
		}
	}
	
	
	public static Font setEstiloFuente(Workbook excel) {
		Font fuente = excel.createFont();
		fuente.setFontHeightInPoints((short) 10);
		fuente.setFontName("Arial");
		fuente.setBold(true);
		fuente.setColor(IndexedColors.WHITE.getIndex());
		return fuente;
	}
	
	public static CellStyle setEstiloPorcentaje(Workbook excel) {
		Font fuente = excel.createFont();
		fuente.setFontHeightInPoints((short) 10);
		fuente.setFontName("Arial");
		fuente.setBold(false);
		
		CellStyle estiloPorcentaje = excel.createCellStyle();
		estiloPorcentaje.setDataFormat(excel.createDataFormat().getFormat("0%"));
		estiloPorcentaje.setWrapText(true);
		estiloPorcentaje.setAlignment(HorizontalAlignment.CENTER);
		estiloPorcentaje.setVerticalAlignment(VerticalAlignment.CENTER);
		estiloPorcentaje.setFont(fuente);
		return estiloPorcentaje;
	}
	
	public static CellStyle setEstiloHeadIzquierda(Workbook excel) {
		Font fuente = excel.createFont();
		fuente.setFontHeightInPoints((short) 10);
		fuente.setFontName("Arial");
		fuente.setBold(true);
		fuente.setColor(IndexedColors.WHITE.getIndex());
		
		CellStyle estiloCeldaIzquierda = excel.createCellStyle();
		estiloCeldaIzquierda.setWrapText(true);
		estiloCeldaIzquierda.setAlignment(HorizontalAlignment.LEFT);
		estiloCeldaIzquierda.setVerticalAlignment(VerticalAlignment.CENTER);
		estiloCeldaIzquierda.setFont(fuente);
		estiloCeldaIzquierda.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
		estiloCeldaIzquierda.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return estiloCeldaIzquierda;
	}
	
	public static CellStyle setEstiloHeadCentrado(Workbook excel) {
		Font fuente = excel.createFont();
		fuente.setFontHeightInPoints((short) 10);
		fuente.setFontName("Arial");
		fuente.setBold(true);
		fuente.setColor(IndexedColors.WHITE.getIndex());
		
		CellStyle estiloCeldaCentrado = excel.createCellStyle();
		estiloCeldaCentrado.setWrapText(true);
		estiloCeldaCentrado.setAlignment(HorizontalAlignment.CENTER);
		estiloCeldaCentrado.setVerticalAlignment(VerticalAlignment.CENTER);
		estiloCeldaCentrado.setFont(fuente);
		estiloCeldaCentrado.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
		estiloCeldaCentrado.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return estiloCeldaCentrado;
	}

	public static CellStyle setEstiloNormalCentrado(Workbook excel) {
		Font fuente = excel.createFont();
		fuente.setFontHeightInPoints((short) 10);
		fuente.setFontName("Arial");
		fuente.setBold(false);

		CellStyle estiloCeldaCentrado = excel.createCellStyle();
		estiloCeldaCentrado.setWrapText(true);
		estiloCeldaCentrado.setAlignment(HorizontalAlignment.CENTER);
		estiloCeldaCentrado.setVerticalAlignment(VerticalAlignment.CENTER);
		estiloCeldaCentrado.setFont(fuente);
		return estiloCeldaCentrado;
	}
	public static CellStyle setEstiloNormalIzquierdo(Workbook excel) {
		Font fuente = excel.createFont();
		fuente.setFontHeightInPoints((short) 10);
		fuente.setFontName("Arial");
		fuente.setBold(false);

		CellStyle estiloCeldaCentrado = excel.createCellStyle();
		estiloCeldaCentrado.setWrapText(true);
		estiloCeldaCentrado.setAlignment(HorizontalAlignment.LEFT);
		estiloCeldaCentrado.setVerticalAlignment(VerticalAlignment.CENTER);
		estiloCeldaCentrado.setFont(fuente);
		return estiloCeldaCentrado;
	}
	
}
