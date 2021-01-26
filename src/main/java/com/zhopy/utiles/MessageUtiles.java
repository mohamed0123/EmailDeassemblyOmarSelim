package com.zhopy.utiles;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hsmf.datatypes.StringChunk;

public class MessageUtiles {

	private static MAPIMessage initMessageObjectFromMsgFile(String filePath) {
		MAPIMessage msg = null;
		try {
			msg = new MAPIMessage(new File(filePath));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return msg;
	}

	private static void closeMessage(MAPIMessage msg) {
		try {
			if (msg != null)
				msg.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void closeOutputStream(OutputStream fileOut) {
		if (fileOut != null) {
			try {
				fileOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	private static boolean writeExcelFile(String tempExcelPath, AttachmentChunks attachmentChunks) {
		File f = new File(tempExcelPath);
		OutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(f);
			fileOut.write(attachmentChunks.getAttachData().getValue());
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {
			closeOutputStream(fileOut);
		}
		return true;
	}

	public static String attachedExcelHandler(String tempExcelPath , String temMsgDir) {
		MAPIMessage msg = null;
		try {
			msg = initMessageObjectFromMsgFile(temMsgDir);
			if (msg == null)
				return "cann't read message file";

			return writeExcelToPath(msg, tempExcelPath);
		} finally {
			closeMessage(msg);
		}
	}

	private static String writeExcelToPath(MAPIMessage msg, String tempExcelPath) {
		AttachmentChunks[] attachments = msg.getAttachmentFiles();
		if (attachments.length > 0) {
			for (AttachmentChunks attachmentChunks : attachments) {
				StringChunk ext = attachmentChunks.getAttachExtension();
				if (ext.getValue().equalsIgnoreCase(".xlsx")) {
					boolean excelStatus = writeExcelFile(tempExcelPath, attachmentChunks);
					if (excelStatus)
						return null;
					else
						return "cann't write Excel file";
				}
			}
		}
		return "cann't write Excel file";
	}

}
