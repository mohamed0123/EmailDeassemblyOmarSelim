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
			e.printStackTrace();
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
			// rowIterator.next(); // Skip Header
			Row headerRow = findHeaderRow(rowIterator);
			InputDto excelDataIndexces = findDataIndexes(headerRow);
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				attachedExcelDtoList.add(fillInputDto(excelDataIndexces, row));
//				attachedExcelDtoList.add(new InputDto(getCellStringValue(row, 0), getCellStringValue(row, 3),
//						getCellStringValue(row, 4), originalEmail));
			}
		}
		return attachedExcelDtoList;
	}

	public static InputDto findDataIndexes(Row row) {

		InputDto excelDataIndexces = new InputDto();

		for (int i = 0; i < row.getRowNum(); i++) {
			String col = getCellStringValue(row, i);
			if (col.equalsIgnoreCase("Manufacturer Part Number (MPN)")) {
				excelDataIndexces.setManufacturerPartNumberMpn(i + "");
			}
			if (col.equalsIgnoreCase("Manufacturer Name")) {
				excelDataIndexces.setManufacturerName(i + "");
			}
			if (col.equalsIgnoreCase("Customer Internal Part Number")) {
				excelDataIndexces.setCustomerInternalPartNumber(i + "");
			}
			if (col.equalsIgnoreCase("Product Description")) {
				excelDataIndexces.setProductDescription(i + "");
			}
			if (col.equalsIgnoreCase("Comment")) {
				excelDataIndexces.setComment(i + "");
			}
			if (col.toLowerCase().contains("Corrected Manufacturer Name".toLowerCase())) {
				excelDataIndexces.setCorrectedManufacturerNameIfNecessary(i + "");
			}
			if (col.toLowerCase().contains("Corrected MPN".toLowerCase())) {
				excelDataIndexces.setCorrectedMpnToBeFilledIfMpnIsIncorrectOrInvalid(i + "");
			}
			if (col.toLowerCase().contains("Lifecycle Status".toLowerCase())) {
				excelDataIndexces.setLifecycleStatus(i + "");
			}
			if (col.toLowerCase().equalsIgnoreCase("LTB date".toLowerCase())) {
				excelDataIndexces.setLtbDateTheLastDateByWhenTheCustomerCanOrderThePart(i + "");
			}
			if (col.equalsIgnoreCase(
					"Reason for NRND/ Discontinued/Obsoleted (If part is NRND/Discontinued/obsoleted)")) {
				excelDataIndexces.setReasonForNrndDiscontinuedObsoletedIfPartIsNrndDiscontinuedObsoleted(i + "");
			}
			if (col.equalsIgnoreCase("Part_Design_Type)")) {
				excelDataIndexces.setPartDesignType(i + "");
			}
			if (col.equalsIgnoreCase("RoHS (2011/65/EU) Status (Select option 1, 2, or 3)")) {
				excelDataIndexces.setRohs(i + "");
			}
			if (col.equalsIgnoreCase("EU RoHS Exemption List (Click on embedded link)")) {
				excelDataIndexces.setEuRohsExemptionListClickOnEmbeddedLink(i + "");
			}
			if (col.equalsIgnoreCase("RoHS (2015/863) New Added 4 phthalates Status (Select option 1-Yes, 2-No)")) {
				excelDataIndexces.setRohs2015863NewAdded4PhthalatesStatusSelectOption1Yes2No(i + "");
			}
			if (col.equalsIgnoreCase("Active RoHS Replacement MPN")) {
				excelDataIndexces.setActiveRohsReplacementMpn(i + "");
			}
			if (col.equalsIgnoreCase("Form Fit Function Compatibility")) {
				excelDataIndexces.setFormFitFunctionCompatibility(i + "");
			}
			if (col.equalsIgnoreCase("EU RoHS Exemption List for Replacement part (Refer to link)")) {
				excelDataIndexces.setEuRohsExemptionListForReplacementPartReferToLink(i + "");
			}
		}
		return excelDataIndexces;
	}

	public static InputDto fillInputDto(InputDto excelDataIndexces, Row row) {
		InputDto dto = new InputDto();
		dto.setActiveRohsReplacementMpn(getCellStringValue(row, excelDataIndexces.getActiveRohsReplacementMpn()));
		dto.setComment(getCellStringValue(row, excelDataIndexces.getComment()));
		dto.setCorrectedManufacturerNameIfNecessary(
				getCellStringValue(row, excelDataIndexces.getCorrectedManufacturerNameIfNecessary()));
		dto.setCorrectedManufacturerNameIfNecessarySeeComments(
				getCellStringValue(row, excelDataIndexces.getCorrectedManufacturerNameIfNecessarySeeComments()));
		dto.setCorrectedMpnToBeFilledIfMpnIsIncorrectOrInvalid(
				getCellStringValue(row, excelDataIndexces.getCorrectedMpnToBeFilledIfMpnIsIncorrectOrInvalid()));
		dto.setCustomerInternalPartNumber(getCellStringValue(row, excelDataIndexces.getCustomerInternalPartNumber()));
		dto.setEuRohsExemptionListClickOnEmbeddedLink(
				getCellStringValue(row, excelDataIndexces.getEuRohsExemptionListClickOnEmbeddedLink()));
		dto.setEuRohsExemptionListForReplacementPartReferToLink(
				getCellStringValue(row, excelDataIndexces.getEuRohsExemptionListForReplacementPartReferToLink()));
		dto.setFormFitFunctionCompatibility(
				getCellStringValue(row, excelDataIndexces.getFormFitFunctionCompatibility()));
		dto.setLifecycleStatus(getCellStringValue(row, excelDataIndexces.getLifecycleStatus()));
		dto.setLtbDateTheLastDateByWhenTheCustomerCanOrderThePart(
				getCellStringValue(row, excelDataIndexces.getLtbDateTheLastDateByWhenTheCustomerCanOrderThePart()));

		dto.setManufacturerName(getCellStringValue(row, excelDataIndexces.getManufacturerName()));
		dto.setManufacturerPartNumberMpn(getCellStringValue(row, excelDataIndexces.getManufacturerPartNumberMpn()));
		dto.setPartDesignType(getCellStringValue(row, excelDataIndexces.getPartDesignType()));
		dto.setProductDescription(getCellStringValue(row, excelDataIndexces.getProductDescription()));
		dto.setReasonForNrndDiscontinuedObsoletedIfPartIsNrndDiscontinuedObsoleted(getCellStringValue(row,
				excelDataIndexces.getReasonForNrndDiscontinuedObsoletedIfPartIsNrndDiscontinuedObsoleted()));
		dto.setRohs(getCellStringValue(row, excelDataIndexces.getRohs()));
		dto.setRohs2015863NewAdded4PhthalatesStatusSelectOption1Yes2No(getCellStringValue(row,
				excelDataIndexces.getRohs2015863NewAdded4PhthalatesStatusSelectOption1Yes2No()));

		return dto;
	}

	public static List<InputDto> readXlsx(String path, String originalEmail) throws Exception {
		List<InputDto> attachedExcelDtoList = new ArrayList<>();
		try (XSSFWorkbook myWorkBook = new XSSFWorkbook(new File(path))) {
			XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			Iterator<Row> rowIterator = mySheet.iterator();
			Row headerRow = findHeaderRow(rowIterator);
			InputDto excelDataIndexces = findDataIndexes(headerRow);
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				attachedExcelDtoList.add(fillInputDto(excelDataIndexces, row));
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
//				private String partNumber;
//				private String description;
//				private String lcStatus;
//				private String originalEmail;
				InputDto inputDto = new InputDto();
				inputDto.setManufacturerPartNumberMpn(getCellStringValue(row, 0));
				inputDto.setProductDescription(getCellStringValue(row, 1));
				inputDto.setLifecycleStatus(getCellStringValue(row, 2));
				inputDto.setOriginalEmail(getCellStringValue(row, 3));
				inputDtoList.add(inputDto);
			}
		}
		return inputDtoList;
	}

	private static Row findHeaderRow(Iterator<Row> rowIterator) {
		Row header = null;
		for (int i = 0; i < 6; i++) {
			header = rowIterator.next();
		}
		return header;
	}

	private static String getCellStringValue(Row row, String index) {
		System.out.println("index : " + index);
		if (index != null)
			return getCellStringValue(row, Integer.parseInt(index));
		else
			return "";
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
