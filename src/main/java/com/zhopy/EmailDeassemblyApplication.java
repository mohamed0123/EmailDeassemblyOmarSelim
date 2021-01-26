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
	public static String logFile = "log";
	public static String logResults = "results";

	public static void main(String[] args) throws Exception {
		init();
		List<InputDto> inputDtoList = ExcelService.loadInput("C:\\Users\\Mohamed_Hamada\\Desktop\\input.xlsx");
		List<String> originalEmailsList = inputDtoList.parallelStream().map(e -> e.getOriginalEmail())
				.filter(e -> e.toLowerCase().endsWith(".msg")).distinct().collect(Collectors.toList());
		originalEmailsList.stream().forEach(e -> mainExecuter(e, filterSpecificEmail(inputDtoList, e)));
	}

	public static List<InputDto> filterSpecificEmail(List<InputDto> inputDtoList, String originalEmail) {
		return inputDtoList.parallelStream().filter(e -> e.getOriginalEmail().equalsIgnoreCase(originalEmail))
				.collect(Collectors.toList());
	}

	public static void init() throws Exception {
		initializeResultFile();
		initializeResultFile();
		Files.createDirectories(Paths.get(msgTempDir));
		Files.createDirectories(Paths.get(excelTempDir));
	}

	public static void initializeResultFile() throws Exception {
		File f = new File(logResults);
		// to remove the old result file
		bwLog = new BufferedWriter(new FileWriter(f, false));
		// to allow append the new results
		bwLog = new BufferedWriter(new FileWriter(f, true));
	}

	public static void initializeLogFile() throws Exception {
		File f = new File(logFile);
		// to remove the old result file
		bwResults = new BufferedWriter(new FileWriter(f, false));
		// to allow append the new results
		bwResults = new BufferedWriter(new FileWriter(f, true));
	}

	public static void writeLogFileHeader() throws Exception {
		bwLog.append("original Mail" + "\t");
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
		List<InputDto> partsStatusResults = ExcelService.loadAttachedExcelDto(tempExcelPath);
		if (partsStatusResults == null) {
			writeLogFile(originalEmail, "Error", "cannot load excel data from path");
			return;
		}
		// remove null or empty fields
		partsStatusResults = partsStatusResults.parallelStream()
				.filter(e -> e.getPartNumber() != null && !e.getPartNumber().isEmpty()).collect(Collectors.toList());
		// write compare method
		comparePartNumbers(originalEmail, partsStatusResults, inputDtoList);

	}

	public static void comparePartNumbers(String originalEmail, List<InputDto> partsStatusResults,
			List<InputDto> inputDtoList) {
		String partStatus;
		if (inputDtoList.size() == partsStatusResults.size()) {
			partStatus = "exact parts";
		} else {

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
}