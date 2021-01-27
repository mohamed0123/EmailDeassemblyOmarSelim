package com.zhopy.utiles;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.zhopy.dto.InputDto;

public class ExcelService {

	public static List<InputDto> loadAttachedExcelDto(String path, String originalEmail) {
		List<InputDto> attachedExcelDtoList = new ArrayList<>();
		try (XSSFWorkbook myWorkBook = new XSSFWorkbook(path)) {
			XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			Iterator<Row> rowIterator = mySheet.iterator();
			skipHeader(rowIterator); // Skip Header
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				attachedExcelDtoList.add(new InputDto(getCellStringValue(row, 0), getCellStringValue(row, 3),
						getCellStringValue(row, 7), originalEmail));
			}
		} catch (IOException e) {
			e.printStackTrace();
			return null;
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
				val = String.valueOf(formatter.formatCellValue(cell));
				break;
			case STRING:
				val = formatter.formatCellValue(cell);
				break;
			default:
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
