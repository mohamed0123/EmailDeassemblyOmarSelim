package com.zhopy.utiles;

import java.io.File;
import java.io.IOException;
import java.net.URL;

import org.apache.commons.io.FileUtils;

public class DownloadUtiles {

	public static boolean downloadUrl(String url, String destination) {
		try {
			if (GeneralUtiles.checkFileExistance(destination)) {
				return true;
			}
			System.out.println(url.trim());

			if (!copyFile(url, destination))
				FileUtils.copyURLToFile(new URL(url.trim()), new File(destination));
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.getMessage());
			return false;
		}
		return true;
	}

	public static boolean copyFile(String url, String destination) {
		try {
			FileUtils.copyFile(new File(GeneralUtiles.getLocalPath(url)), new File(destination));
			return true;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		}
	}

}
