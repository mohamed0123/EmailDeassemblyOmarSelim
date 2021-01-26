package com.zhopy.utiles;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URL;

import org.apache.commons.compress.utils.IOUtils;

public class DownloadUtiles {

	public static boolean downloadUrl(String url, String savedPath) {
		try {
			if (GeneralUtiles.checkFileExistance(savedPath)) {
				return true;
			}
			InputStream inputStream = new URL(url).openStream();
			FileOutputStream fileOS = new FileOutputStream(savedPath);
			long i = IOUtils.copy(inputStream, fileOS, 1024);
			System.out.println("file size " + i);
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.getMessage());
			return false;
		}
		return true;
	}

	
}
