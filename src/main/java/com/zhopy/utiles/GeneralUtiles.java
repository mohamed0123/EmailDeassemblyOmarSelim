package com.zhopy.utiles;

import java.io.File;
import java.util.List;

import org.apache.commons.compress.utils.FileNameUtils;

import net.rationalminds.LocalDateModel;
import net.rationalminds.Parser;

public class GeneralUtiles {

	private static String ExtractDateFromString(String text) {
		Parser parser = new Parser();
		List<LocalDateModel> dates = parser.parse(removeFirstSection(text));
		return dates.get(0).getDateTimeString();
	}

	private static String removeFirstSection(String path) {
		return path.replaceAll("http://download.siliconexpert.com/pdfs2/", "")
				.replaceAll("https://download.siliconexpert.com/pdfs/", "");
	}

	
	public static String getLocalPath(String path) {
		return path.replaceAll("http://download.siliconexpert.com/pdfs", "//10.0.1.112/pdfs")
				.replaceAll("https://download.siliconexpert.com/pdfs", "//10.0.1.112/pdfs");
	}

	public static String generateTempMessageFile(String msgUrl, String tempDir) {
		String date = ExtractDateFromString(msgUrl);
		String fileName = FileNameUtils.getBaseName(msgUrl);
		String fileExtention = FileNameUtils.getExtension(msgUrl);
		return tempDir + date + "_" + fileName + "_" + fileExtention;
	}

	public static String generateTempExcelFile(String msgUrl, String tempDir) {
		String date = ExtractDateFromString(msgUrl);
		String fileName = removeWhiteSpacesFromString(FileNameUtils.getBaseName(msgUrl));
		return tempDir + date + "_" + fileName + "Excel_.xlsx";
	}

	public static boolean checkFileExistance(String filePath) {
		File f = new File(filePath);
		if (f.exists() && f.length() > 0)
			return true;
		return false;
	}
	
	public static String removeWhiteSpacesFromString(String token) {
		if (token == null)
			return "";
		return token.replaceAll("\\t+", " ").replaceAll("\\r+", " ").replaceAll("\\n+", " ").replaceAll("\\r\\n+", " ");
	}
}
