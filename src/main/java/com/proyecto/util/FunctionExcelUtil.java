package com.proyecto.util;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class FunctionExcelUtil {

	public static Map<String, Object> validaCabecera(String[] columnas, Sheet sheet) {
		Map<String, Object> salida = new HashMap<>();
		for (Cell cell : sheet.getRow(0)) {
			switch (cell.getCellType()) {
			case BLANK:
				salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1)	+ ") es de tipo Texto (No vacio)");
				return salida;
			case NUMERIC:
				salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1)	+ ") es de tipo Texto (No numérico)");
				return salida;
			case STRING:
				if (!columnas[cell.getColumnIndex()].equals(cell.getStringCellValue())) {
					salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1) + ") debe tener la columna : " + columnas[cell.getColumnIndex()]);
					return salida;
				}
				break;
			default:
				salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1) + ") es de tipo Texto");
			}
		}
		return salida;
	}
	
	public static Map<String, Object> validaCabecera(ArrayList<String> columnas, Sheet sheet) {
		Map<String, Object> salida = new HashMap<>();
		for (Cell cell : sheet.getRow(0)) {
			switch (cell.getCellType()) {
			case BLANK:
				salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1)	+ ") es de tipo Texto (No vacio)");
				return salida;
			case NUMERIC:
				salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1)	+ ") es de tipo Texto (No numérico)");
				return salida;
			case STRING:
				if (!columnas.get(cell.getColumnIndex()).equals(cell.getStringCellValue())) {
					salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1) + ") debe tener la columna : " + columnas.get(cell.getColumnIndex()));
					return salida;
				}
				break;
			default:
				salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1) + ") es de tipo Texto");
			}
		}
		return salida;
	}

	public static Map<String, Object> validaTipoDatos(CellType[] tipos, Sheet sheet) {
		Map<String, Object> salida = new HashMap<>();
		boolean noEsPrimero = false;
		for (Row row : sheet) {
			if (noEsPrimero) {
				for (Cell cell : row) {
					if (!tipos[cell.getColumnIndex()].equals(cell.getCellType())) {
						salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1) + ") debe tener formato : " + tipoCadena(tipos[cell.getColumnIndex()]));
						return salida;
					}
				}
			}
			noEsPrimero = true;
		}
		return salida;
	}
	
	public static Map<String, Object> validaTipoDatos(ArrayList<CellType> tipos, Sheet sheet) {
		Map<String, Object> salida = new HashMap<>();
		boolean noEsPrimero = false;
		for (Row row : sheet) {
			if (noEsPrimero) {
				for (Cell cell : row) {
					if (!tipos.get(cell.getColumnIndex()).equals(cell.getCellType())) {
						salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1) + ") debe tener formato : " + tipoCadena(tipos.get(cell.getColumnIndex())) );
						return salida;
					}
				}
			}
			noEsPrimero = true;
		}
		return salida;
	}
	

	public static Map<String, Object> validaTipoDatosArreglo(CellType[][] tipos, Sheet sheet) {
		Map<String, Object> salida = new HashMap<>();
		boolean noEsPrimero = false;
		for (Row row : sheet) {
			if (noEsPrimero) {
				for (Cell cell : row) {
					if (noExisteTipoArray(tipos[cell.getColumnIndex()],cell.getCellType())) {
						salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1) + ") debe tener formato : " + tipoCadenaArray(tipos[cell.getColumnIndex()]));
						return salida;
					}
				}
			}
			noEsPrimero = true;
		}
		return salida;
	}
	
	public static Map<String, Object> validaTipoDatosArrayList(ArrayList<ArrayList<CellType>> tipos, Sheet sheet) {
		Map<String, Object> salida = new HashMap<>();
		boolean noEsPrimero = false;
		for (Row row : sheet) {
			if (noEsPrimero) {
				for (Cell cell : row) {
					if (noExisteTipoArrayList(tipos.get(cell.getColumnIndex()),cell.getCellType())) {
						salida.put("mensaje", "Celda (" + (cell.getRowIndex() + 1) + "," + (cell.getColumnIndex() + 1) + ") debe tener formato : " + tipoCadenaArrayList(tipos.get(cell.getColumnIndex())));
						return salida;
					}
				}
			}
			noEsPrimero = true;
		}
		return salida;
	}
	
	public static boolean noExisteTipoArrayList(ArrayList<CellType> lista, CellType aux) {
		for (CellType cellType : lista) {
			if (cellType.equals(aux)) {
				return false;
			}
		} 
		return true;
	}
	
	public static boolean noExisteTipoArray(CellType[] array, CellType aux) {
		for (CellType cellType : array) {
			if (cellType.equals(aux)) {
				return false;
			}
		} 
		return true;
	}
	
	public static String tipoCadenaArray(CellType[] array) {
		String cadena = "";
		ArrayList<String> lista = new ArrayList<String>();
		for (CellType x : array) {
			switch (x) {
				case FORMULA: lista.add("FÓRMULA");	break;
				case BLANK: lista.add("BLANCO");	break;
				case BOOLEAN: lista.add("BOOLEANO");	break;
				case NUMERIC: lista.add("NUMÉRICO");	break;
				case STRING: lista.add("ALFANUMÉRICO");	break;
				default: lista.add("NINGUNO");	break;
			}
		} 
		
		for (String aux : lista) {
			cadena += aux;
			if (lista.indexOf(aux) != lista.size()-1) {
				cadena += ", ";
			}
		}
		cadena +=".";
		return cadena;
	}
	
	public static String tipoCadenaArrayList(ArrayList<CellType> array) {
		String cadena = "";
		ArrayList<String> lista = new ArrayList<String>();
		for (CellType x : array) {
			switch (x) {
				case FORMULA: lista.add("FÓRMULA");	break;
				case BLANK: lista.add("BLANCO");	break;
				case BOOLEAN: lista.add("BOOLEANO");	break;
				case NUMERIC: lista.add("NUMÉRICO");	break;
				case STRING: lista.add("ALFANUMÉRICO");	break;
				default: lista.add("NINGUNO");	break;
			}
		} 
		
		for (String aux : lista) {
			cadena += aux;
			if (lista.indexOf(aux) != lista.size()-1) {
				cadena += ", ";
			}
		}
		cadena +=".";
		return cadena;
	}
	
	public static String tipoCadena(CellType aux) {
		String cadena = "";
		switch (aux) {
				case FORMULA: cadena += "FÓRMULA";	break;
				case BLANK: cadena +="BLANCO";	break;
				case BOOLEAN: cadena +="BOOLEAN";	break;
				case NUMERIC: cadena +="NUMÉRICO";	break;
				case STRING: cadena +="ALFANUMÉRICO";	break;
				default: cadena +="NINGUNO";	break;
		} 
		
		cadena +=".";
		return cadena;
	}
	
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
	
	public static String verResultado(int nota) {
		if (nota >= 17) 
			return "Adecuado";
		else if (nota >= 15) 
			return "Medianamente adecuado";
		else 
			return "Con dificultad";
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
	
	public static CellStyle setEstiloHeadIzquierdaInformacion(Workbook excel) {
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
		estiloCeldaIzquierda.setBorderBottom(BorderStyle.THIN);
		estiloCeldaIzquierda.setBorderTop(BorderStyle.THIN);
		estiloCeldaIzquierda.setBorderRight(BorderStyle.THIN);
		estiloCeldaIzquierda.setBorderLeft(BorderStyle.THIN);
		return estiloCeldaIzquierda;
	}
	
	public static CellStyle setEstiloHeadCentradoInformacion(Workbook excel) {
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
		estiloCeldaCentrado.setBorderBottom(BorderStyle.THIN);
		estiloCeldaCentrado.setBorderTop(BorderStyle.THIN);
		estiloCeldaCentrado.setBorderRight(BorderStyle.THIN);
		estiloCeldaCentrado.setBorderLeft(BorderStyle.THIN);
		return estiloCeldaCentrado;
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
	public static CellStyle setEstiloHeadIzquierdaData(Workbook excel) {
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
		estiloCeldaIzquierda.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		estiloCeldaIzquierda.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		estiloCeldaIzquierda.setBorderBottom(BorderStyle.THIN);
		estiloCeldaIzquierda.setBorderTop(BorderStyle.THIN);
		estiloCeldaIzquierda.setBorderRight(BorderStyle.THIN);
		estiloCeldaIzquierda.setBorderLeft(BorderStyle.THIN);
		return estiloCeldaIzquierda;
	}
	
	public static CellStyle setEstiloHeadCentradoData(Workbook excel) {
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
		estiloCeldaCentrado.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		estiloCeldaCentrado.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		estiloCeldaCentrado.setBorderBottom(BorderStyle.THIN);
		estiloCeldaCentrado.setBorderTop(BorderStyle.THIN);
		estiloCeldaCentrado.setBorderRight(BorderStyle.THIN);
		estiloCeldaCentrado.setBorderLeft(BorderStyle.THIN);
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
