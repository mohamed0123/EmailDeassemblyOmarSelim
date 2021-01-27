package com.zhopy.utiles;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.zhopy.dto.InputDto;

public class ExcelService {

	public static List<InputDto> loadAttachedExcelDto(String path, String originalEmail) {
		System.out.println(path);
		List<InputDto> attachedExcelDtoList = new ArrayList<>();
		try {
			attachedExcelDtoList = readXlsx(path, originalEmail);
		} catch (Exception e) {
			try {
				if (e.getMessage().contains(
						"The supplied data appears to be in the OLE2 Format. You are calling the part of POI that deals with OOXML")) {

					attachedExcelDtoList = readXls(path, originalEmail);
					return attachedExcelDtoList;
				}

			} catch (Exception e2) {
				e2.printStackTrace();
			}

			return null;
		}
		return attachedExcelDtoList;
	}

	public static List<InputDto> readXls(String path, String originalEmail) throws Exception {

		File f = new File(path);
		path = path.replace(".xlsx", "xls");
		f.renameTo(new File(path));

		List<InputDto> attachedExcelDtoList = new ArrayList<>();
		InputStream input = new FileInputStream(path);
		POIFSFileSystem fs = new POIFSFileSystem(input); // class permetant de lire fichier
		try (HSSFWorkbook myWorkBook = new HSSFWorkbook(fs)) {
			HSSFSheet mySheet = myWorkBook.getSheetAt(0);
			Iterator<Row> rowIterator = mySheet.iterator();
			rowIterator.next(); // Skip Header
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				attachedExcelDtoList.add(new InputDto(getCellStringValue(row, 0), getCellStringValue(row, 3),
						getCellStringValue(row, 4), originalEmail));
			}
		}
		return attachedExcelDtoList;
	}

	public static List<InputDto> readXlsx(String path, String originalEmail) throws Exception {
		List<InputDto> attachedExcelDtoList = new ArrayList<>();
		try (XSSFWorkbook myWorkBook = new XSSFWorkbook(new File(path))) {
			XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			Iterator<Row> rowIterator = mySheet.iterator();
			skipHeader(rowIterator); // Skip Header
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				attachedExcelDtoList.add(new InputDto(getCellStringValue(row, 0), getCellStringValue(row, 3),
						getCellStringValue(row, 7), originalEmail));
			}
		}
		return attachedExcelDtoList;
	}

	public static List<InputDto> loadInput(String path) throws Exception {
		List<InputDto> inputDtoList = new ArrayList<>();
		try (XSSFWorkbook myWorkBook = new XSSFWorkbook(path)) {
			XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			Iterator<Row> rowIterator = mySheet.iterator();
			rowIterator.next(); // Skip Header
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				inputDtoList.add(new InputDto(getCellStringValue(row, 0), getCellStringValue(row, 1),
						getCellStringValue(row, 2), getCellStringValue(row, 3)));
			}
		}
		return inputDtoList;
	}

	private static void skipHeader(Iterator<Row> rowIterator) {
		for (int i = 0; i < 6; i++) {
			rowIterator.next();
		}
	}

	private static String getCellStringValue(Row row, int index) {
		try {
			String val = "";
			DataFormatter formatter = new DataFormatter();

			Cell cell = row.getCell(index);

			switch (cell.getCellType()) {
			case NUMERIC:
				cell.setCellType(CellType.STRING);
				val = String.valueOf(formatter.formatCellValue(cell));
				break;
			case STRING:
				val = formatter.formatCellValue(cell);
				break;
			default:
				cell.setCellType(CellType.STRING);
				val = String.valueOf(formatter.formatCellValue(cell));
				break;
			}

			return val;
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return "";
	}
}
