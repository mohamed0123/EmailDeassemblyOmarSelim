package com.zhopy;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

import com.zhopy.dto.InputDto;
import com.zhopy.utiles.DownloadUtiles;
import com.zhopy.utiles.ExcelService;
import com.zhopy.utiles.GeneralUtiles;
import com.zhopy.utiles.MessageUtiles;

public class EmailDeassemblyApplication {
	public static BufferedWriter bwLog;
	public static BufferedWriter bwResults;
	public static String msgTempDir = "./msgTemp/";
	public static String excelTempDir = "./excelTemp/";
	public static int seq = 1;
	public static String logFile = "log.txt";
	public static String logResults = "results.txt";

	public static void main(String[] args) throws Exception {

		if (args.length == 0)
			args = new String[] { "C:\\Users\\Mohamed_Hamada\\Desktop\\input.xlsx" };

		init();
		List<InputDto> inputDtoList = ExcelService.loadInput(args[0]);
		List<String> originalEmailsList = inputDtoList.parallelStream().map(e -> e.getOriginalEmail())
				.filter(e -> e != null && e.toLowerCase().endsWith(".msg")).distinct().collect(Collectors.toList());
		originalEmailsList.parallelStream().forEach(e -> mainExecuter(e, filterSpecificEmail(inputDtoList, e)));
		flush();
		closeBw();
	}

	public static void closeBw() {
		try {
			bwLog.close();
			bwResults.close();
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static List<InputDto> filterSpecificEmail(List<InputDto> inputDtoList, String originalEmail) {
		return inputDtoList.parallelStream().filter(e -> e.getOriginalEmail().equalsIgnoreCase(originalEmail))
				.distinct().collect(Collectors.toList());
	}

	public static void init() throws Exception {
		initializeResultFile();
		initializeLogFile();
		Files.createDirectories(Paths.get(msgTempDir));
		Files.createDirectories(Paths.get(excelTempDir));
		writeResultsFileHeader();
		writeLogFileHeader();
	}

	public static void initializeResultFile() throws Exception {
		File f = new File(logResults);
		// to remove the old result file
		bwResults = new BufferedWriter(new FileWriter(f, false));
		// to allow append the new results
		bwResults = new BufferedWriter(new FileWriter(f, true));

	}

	public static void initializeLogFile() throws Exception {
		File f = new File(logFile);
		// to remove the old result file
		bwLog = new BufferedWriter(new FileWriter(f, false));
		// to allow append the new results
		bwLog = new BufferedWriter(new FileWriter(f, true));
	}

	public static void writeLogFileHeader() throws Exception {
		bwLog.write("original Mail" + "\t");
		bwLog.append("Status" + "\t");
		bwLog.append("message" + "\r\n");
	}

	public static synchronized void writeLogFile(String originalEmail, String status, String message) {
		try {
			bwLog.append(originalEmail + "\t");
			bwLog.append(status + "\t");
			bwLog.append(message + "\r\n");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void mainExecuter(String originalEmail, List<InputDto> inputDtoList) {
		String temMsgDir = GeneralUtiles.generateTempMessageFile(originalEmail, msgTempDir);
		boolean isFileDownloaded = DownloadUtiles.downloadUrl(originalEmail, temMsgDir);
		if (!isFileDownloaded) {
			writeLogFile(originalEmail, "Error", "cann't specified file path");
			return;
		}
		String tempExcelPath = GeneralUtiles.generateTempExcelFile(originalEmail, excelTempDir);
		String errorMessage = MessageUtiles.attachedExcelHandler(tempExcelPath, temMsgDir);
		if (errorMessage != null) {
			writeLogFile(originalEmail, "Error", errorMessage);
			return;
		}
		List<InputDto> partsStatusResults = ExcelService.loadAttachedExcelDto(tempExcelPath, originalEmail);
		if (partsStatusResults == null) {
			writeLogFile(originalEmail, "Error", "cannot load excel data from path");
			return;
		}
		// remove null or empty fields
		partsStatusResults = partsStatusResults.parallelStream()
				.filter(e -> e.getPartNumber() != null && !e.getPartNumber().isEmpty()).collect(Collectors.toList());
		// write compare method
		comparePartNumbers(originalEmail, partsStatusResults, inputDtoList);
		flush();

	}

	public static void flush() {
		try {
			bwResults.flush();
			bwLog.flush();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static List<InputDto> getSharedList(List<InputDto> listOne, List<InputDto> listTwo) {
		List<InputDto> listOneList = listOne.stream()
				.filter(two -> listTwo.stream().anyMatch(one -> one.getPartNumber().equals(two.getPartNumber())))
				.collect(Collectors.toList());
		return listOneList;
	}

	public static List<InputDto> getDiffrenceList(List<InputDto> listOne, List<InputDto> listTwo) {
		Set<String> partNumbersList = listTwo.stream().map(InputDto::getPartNumber).collect(Collectors.toSet());
		List<InputDto> diff = listOne.stream().filter(inputDto -> !partNumbersList.contains(inputDto.getPartNumber()))
				.collect(Collectors.toList());
		return diff;
	}

	public static void writeAddedDeletedToResultsFile(List<InputDto> records, String status) {
		records.stream().forEach(e -> {
			try {
				bwResults.append(e.getPartNumber() + "\t");
				bwResults.append(e.getDescription() + "\t");
				bwResults.append(e.getLcStatus() + "\t");
				bwResults.append(e.getOriginalEmail() + "\t");
				bwResults.append(status + "\t");
				bwResults.append("" + "\r\n");
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		});
	}

	public static void comparePartNumbers(String originalEmail, List<InputDto> partsStatusResults,
			List<InputDto> inputDtoList) {

		List<InputDto> added = getDiffrenceList(partsStatusResults, inputDtoList);
		List<InputDto> deleted = getDiffrenceList(inputDtoList, partsStatusResults);
		List<InputDto> common = getSharedList(inputDtoList, partsStatusResults);

		writeAddedDeletedToResultsFile(added, "Extra Added Parts");
		writeAddedDeletedToResultsFile(deleted, "Missed Parts");
		commonToResultsHandler(originalEmail, common, deleted, added, partsStatusResults);
		System.out.println("<<<<<<<<<<<>>>>>>>>>>>");

	}

	public static void commonToResultsHandler(String originalEmail, List<InputDto> common, List<InputDto> deleted,
			List<InputDto> added, List<InputDto> partsStatusResults) {
		String partsStatus;
		if (added.size() != 0 || deleted.size() != 0)
			partsStatus = "Miss Match Parts";
		else
			partsStatus = "No Changed Parts";
		try {
			boolean isLcStatusChanged = writeCommonToResultsFile(common, partsStatusResults);
			if (isLcStatusChanged)
				writeLogFile(originalEmail, partsStatus + "|LC Have Change", "");
			else
				writeLogFile(originalEmail, partsStatus + "|No Change In LC", "cannot load excel data from path");
		} catch (Exception e) {
			e.printStackTrace();
			writeLogFile(originalEmail, partsStatus + "|LC Have Change", "cannot write to reslts file");
		}
	}

	public static boolean writeCommonToResultsFile(List<InputDto> common, List<InputDto> partsStatusResults)
			throws Exception {
		boolean isLCStatusHaveChange = false;

		for (InputDto e : partsStatusResults) {
			String[] statuLcStatus = lcStatus(e, partsStatusResults);
			bwResults.append(e.getPartNumber() + "\t");
			bwResults.append(e.getDescription() + "\t");
			bwResults.append(e.getLcStatus() + "\t");
			bwResults.append(e.getOriginalEmail() + "\t");
			bwResults.append(statuLcStatus[0] + "\t");
			bwResults.append(statuLcStatus[1] + "\r\n");

			if (statuLcStatus[0].equalsIgnoreCase("No Change"))
				isLCStatusHaveChange = true;
		}
		return isLCStatusHaveChange;

	}

	public static void writeResultsFileHeader() throws Exception {
		bwResults.write("Part Number" + "\t");
		bwResults.append("Description" + "\t");
		bwResults.append("LC Status" + "\t");
		bwResults.append("original Mail" + "\t");
		bwResults.append("Status" + "\t");
		bwResults.append("attached LC Status" + "\r\n");

	}

	public static String[] lcStatus(InputDto e, List<InputDto> partsStatusResults) {
		InputDto attachedRecord = partsStatusResults.parallelStream()
				.filter(c -> c.getPartNumber() == e.getPartNumber()).findFirst().orElse(null);
		if (attachedRecord.getLcStatus() != e.getLcStatus()) {
			return new String[] { "LC Changed", attachedRecord.getLcStatus() };
		} else {
			return new String[] { "No Change", "" };
		}
	}

}
