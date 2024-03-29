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
import com.zhopy.dto.MessageDto;
import com.zhopy.utiles.DownloadUtiles;
import com.zhopy.utiles.ExcelService;
import com.zhopy.utiles.GeneralUtiles;
import com.zhopy.utiles.MessageUtiles;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class EmailDeassemblyApplication {
	public static BufferedWriter bwLog;
	public static BufferedWriter bwResults;
	public static String msgTempDir = "./msgTemp/";
	public static String excelTempDir = "./excelTemp/";
	public static int seq = 1;
	public static String logFile = "log.txt";
	public static String logResults = "results.txt";

	public static void main(String[] args) throws Exception {

		log.info("<<<<<<<<<<<<<<<<<<<Test>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
		if (args.length == 0)
			args = new String[] { "C:\\Users\\wolfn\\Downloads\\input.xlsx" };

		init();
		List<InputDto> inputDtoList = ExcelService.loadInput(args[0]);
		List<String> originalEmailsList = inputDtoList.parallelStream().map(e -> e.getOriginalEmail())
				.filter(e -> e != null && e.toLowerCase().endsWith(".msg")).distinct().collect(Collectors.toList());
		originalEmailsList.stream().forEach(e -> mainExecuter(e, filterSpecificEmail(inputDtoList, e)));
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
		MessageDto messageDto = MessageUtiles.attachedExcelHandler(tempExcelPath, temMsgDir);
		if (messageDto.getErrMsg() != null) {
			writeLogFile(originalEmail, "Error", messageDto.getErrMsg());
			return;
		}
		List<InputDto> partsStatusResults = ExcelService.loadAttachedExcelDto(tempExcelPath, originalEmail);
		if (partsStatusResults == null) {
			writeLogFile(originalEmail, "Error", "cannot load excel data from path");
			return;
		}
		// remove null or empty fields
		partsStatusResults = partsStatusResults.parallelStream()
				.filter(e -> e.getManufacturerPartNumberMpn() != null && !e.getManufacturerPartNumberMpn().isEmpty())
				.collect(Collectors.toList());
		// write compare method
		comparePartNumbers(originalEmail, partsStatusResults, inputDtoList, messageDto);
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
				.filter(two -> listTwo.stream()
						.anyMatch(one -> one.getManufacturerPartNumberMpn().equals(two.getManufacturerPartNumberMpn())))
				.collect(Collectors.toList());
		return listOneList;
	}

	private static String cleanCell(String cell) {
		if (cell != null)
			return cell.replace("\n", " ").replace("\"", " ").replace("\r", " ").replace("\r\n", " ").replace("\t",
					" ");
		return " ";
	}

	public static List<InputDto> getDiffrenceList(List<InputDto> listOne, List<InputDto> listTwo) {
		Set<String> partNumbersList = listTwo.stream().map(InputDto::getManufacturerPartNumberMpn)
				.collect(Collectors.toSet());
		List<InputDto> diff = listOne.stream()
				.filter(inputDto -> !partNumbersList.contains(inputDto.getManufacturerPartNumberMpn()))
				.collect(Collectors.toList());
		return diff;
	}

	public static void writeAddedDeletedToResultsFile(List<InputDto> records, String status, MessageDto messageDto,
			String originalEmail) {
		records.stream().forEach(e -> {

			if (e.getManufacturerPartNumberMpn().equalsIgnoreCase("CRCW040215K4FKED")) {
				System.out.println("next is error");
			}

			try {
				bwResults.append(cleanCell(e.getManufacturerPartNumberMpn()) + "\t");
				bwResults.append(cleanCell(e.getProductDescription()) + "\t");
				bwResults.append(cleanCell(e.getLifecycleStatus()) + "\t");
				bwResults.append(cleanCell(originalEmail) + "\t");
				bwResults.append(cleanCell(status) + "\t");
				bwResults.append("\t");
				bwResults.append(cleanCell(messageDto.getCc()) + "\t");
				bwResults.append(cleanCell(messageDto.getFrom()) + "\t");
				bwResults.append(cleanCell(messageDto.getRecivedDate()) + "\t");
				bwResults.append(cleanCell(messageDto.getTo()) + "\t");
				bwResults.append(cleanCell(e.getManufacturerName()) + "\t");
				bwResults.append(cleanCell(e.getCustomerInternalPartNumber()) + "\t");
				bwResults.append(cleanCell(e.getComment()) + "\t");
				bwResults.append(cleanCell(e.getCorrectedManufacturerNameIfNecessary()) + "\t");
				bwResults.append(cleanCell(e.getCorrectedMpnToBeFilledIfMpnIsIncorrectOrInvalid()) + "\t");
				bwResults.append(cleanCell(e.getLtbDateTheLastDateByWhenTheCustomerCanOrderThePart()) + "\t");
				bwResults.append(
						cleanCell(e.getReasonForNrndDiscontinuedObsoletedIfPartIsNrndDiscontinuedObsoleted()) + "\t");
				bwResults.append(cleanCell(e.getPartDesignType()) + "\t");
				bwResults.append(cleanCell(e.getRohs()) + "\t");
				bwResults.append(cleanCell(e.getEuRohsExemptionListClickOnEmbeddedLink()) + "\t");
				bwResults.append(cleanCell(e.getRohs2015863NewAdded4PhthalatesStatusSelectOption1Yes2No()) + "\t");
				bwResults.append(cleanCell(e.getActiveRohsReplacementMpn()) + "\t");
				bwResults.append(cleanCell(e.getFormFitFunctionCompatibility()) + "\t");
				bwResults.append(cleanCell(e.getEuRohsExemptionListForReplacementPartReferToLink()) + "\t");

				bwResults.append("" + "\r\n");
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		});
	}

	public static void comparePartNumbers(String originalEmail, List<InputDto> partsStatusResults,
			List<InputDto> inputDtoList, MessageDto messageDto) {

		List<InputDto> added = getDiffrenceList(partsStatusResults, inputDtoList);
		List<InputDto> deleted = getDiffrenceList(inputDtoList, partsStatusResults);
		List<InputDto> common = getSharedList(inputDtoList, partsStatusResults);

		writeAddedDeletedToResultsFile(added, "Extra Added Parts", messageDto, originalEmail);
		writeAddedDeletedToResultsFile(deleted, "Missed Parts", messageDto, originalEmail);
		commonToResultsHandler(originalEmail, common, deleted, added, partsStatusResults, messageDto);
		System.out.println("<<<<<<<<<<<>>>>>>>>>>>");

	}

	public static void commonToResultsHandler(String originalEmail, List<InputDto> common, List<InputDto> deleted,
			List<InputDto> added, List<InputDto> partsStatusResults, MessageDto messageDto) {
		String partsStatus;
		if (added.size() != 0 || deleted.size() != 0)
			partsStatus = "Miss Match Parts";
		else
			partsStatus = "No Changed Parts";
		try {
			boolean isLcStatusChanged = writeCommonToResultsFile(common, partsStatusResults, originalEmail);
			if (isLcStatusChanged)
				writeLogFile(originalEmail, partsStatus + "|LC Have Change", "");
			else
				writeLogFile(originalEmail, partsStatus + "|No Change In LC", "");
		} catch (Exception e) {
			e.printStackTrace();
			writeLogFile(originalEmail, partsStatus + "|LC Have Change", "cannot write to reslts file");
		}
	}

	public static boolean writeCommonToResultsFile(List<InputDto> common, List<InputDto> partsStatusResults,
			String originalEmail) throws Exception {
		boolean isLCStatusHaveChange = false;

		for (InputDto e : partsStatusResults) {
			String[] statuLcStatus = lcStatus(e, partsStatusResults);
			bwResults.append(cleanCell(e.getManufacturerPartNumberMpn()) + "\t");
			bwResults.append(cleanCell(e.getProductDescription()) + "\t");
			bwResults.append(cleanCell(e.getLifecycleStatus()) + "\t");
			bwResults.append(cleanCell(originalEmail) + "\t");
			bwResults.append(cleanCell(statuLcStatus[0]) + "\t");
			bwResults.append(cleanCell(statuLcStatus[1]) + "\r\n");

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
		bwResults.append("attached LC Status" + "\t");
		bwResults.append("cc" + "\t");
		bwResults.append("from" + "\t");
		bwResults.append("RecivedDate" + "\t");
		bwResults.append("to" + "\t");
		bwResults.append("ManufacturerName" + "\t");
		bwResults.append("CustomerInternalPartNumber" + "\t");
		bwResults.append("Comment" + "\t");
		bwResults.append("CorrectedManufacturerNameIfNecessary" + "\t");
		bwResults.append("CorrectedMpnToBeFilledIfMpnIsIncorrectOrInvalid" + "\t");
		bwResults.append("LtbDateTheLastDateByWhenTheCustomerCanOrderThePart" + "\t");
		bwResults.append("ReasonForNrndDiscontinuedObsoletedIfPartIsNrndDiscontinuedObsoleted" + "\t");
		bwResults.append("PartDesignType" + "\t");
		bwResults.append("Rohs" + "\t");
		bwResults.append("EuRohsExemptionListClickOnEmbeddedLink" + "\t");
		bwResults.append("Rohs2015863NewAdded4PhthalatesStatusSelectOption1Yes2No" + "\t");
		bwResults.append("ActiveRohsReplacementMpn" + "\t");
		bwResults.append("FormFitFunctionCompatibility" + "\t");
		bwResults.append("EuRohsExemptionListForReplacementPartReferToLink" + "\t");

		bwResults.append("\r\n");

	}

	public static String[] lcStatus(InputDto e, List<InputDto> partsStatusResults) {
		InputDto attachedRecord = partsStatusResults.parallelStream()
				.filter(c -> c.getManufacturerPartNumberMpn() == e.getManufacturerPartNumberMpn()).findFirst()
				.orElse(null);
		if (attachedRecord.getLifecycleStatus() != e.getLifecycleStatus()) {
			return new String[] { "LC Changed", attachedRecord.getLifecycleStatus() };
		} else {
			return new String[] { "No Change", "" };
		}
	}

}
