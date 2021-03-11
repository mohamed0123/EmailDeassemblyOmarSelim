package com.zhopy.utiles;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hsmf.datatypes.Chunks;
import org.apache.poi.hsmf.datatypes.StringChunk;

import com.zhopy.dto.MessageDto;

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

	public static String patternMatchEmails(String str, String type) {
		String emails = null;
		try {
			Pattern pattern = Pattern.compile("^" + type + ":.*\\r\\n", Pattern.MULTILINE);
			Matcher matcher = pattern.matcher(str);
			if (matcher.find()) {
				emails = matcher.group().replace("<", "").replace(">", ";");
			}
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		return emails;
	}

	public static String getStringChunk(MAPIMessage msg) {
		Chunks mainChunks = msg.getMainChunks();
		StringChunk stringChunk = mainChunks.getMessageHeaders();
		String stringChunkValue = stringChunk.getValue();
		return stringChunkValue;
	}

	public static MessageDto attachedExcelHandler(String tempExcelPath, String temMsgDir) {

		MAPIMessage msg = null;
		MessageDto messageDto = new MessageDto();
		try {
			msg = initMessageObjectFromMsgFile(temMsgDir);
			if (msg == null) {
				messageDto.setStatus("Error");
				messageDto.setErrMsg("cann't read message file");
				return messageDto;
			}
			msg.setReturnNullOnMissingChunk(true);
			String stringChunk = getStringChunk(msg);
			messageDto.setTo(msg.getRecipientEmailAddress());
			messageDto.setCc(patternMatchEmails(stringChunk, "CC"));
			messageDto.setFrom(patternMatchEmails(stringChunk, "From"));
			messageDto.setSubject(msg.getSubject());
			SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy HH:mm:ss.SSS");
			messageDto.setRecivedDate(dateFormat.format(msg.getMessageDate().getTime()));

			String writeExcel = writeExcelToPath(msg, tempExcelPath);

			if (writeExcel == null) {
				messageDto.setStatus("Done");
			} else {
				messageDto.setStatus("Error");
				messageDto.setErrMsg(writeExcel);
			}
			return messageDto;
		} catch (Exception e) {
			e.printStackTrace();
			messageDto.setStatus("Error");
			messageDto.setErrMsg(e.getMessage());
			return messageDto;
		} finally {
			closeMessage(msg);
		}
	}

	public static MessageDto attachedExcelHandlerOld(String tempExcelPath, String temMsgDir) {

		MAPIMessage msg = null;
		MessageDto messageDto = new MessageDto();
		try {
			msg = initMessageObjectFromMsgFile(temMsgDir);
			if (msg == null) {
				messageDto.setStatus("Error");
				messageDto.setErrMsg("cann't read message file");
				return messageDto;
			}

			System.out.println(msg.getSummaryInformation());
			System.out.println(msg.getDisplayCC());
			System.out.println(msg.getRecipientNames());
			System.out.println(msg.getHeaders());
			System.out.println(msg.getRecipientEmailAddress());
			System.out.println(msg.getSummaryInformation());
			System.out.println(msg.getSummaryInformation());

			messageDto.setCc(msg.getDisplayCC());
			messageDto.setFrom(msg.getDisplayFrom());
			messageDto.setTo(msg.getRecipientEmailAddress());

			DateTimeFormatter formmat1 = DateTimeFormatter.ofPattern("yyyy-MM-dd", Locale.ENGLISH);

			messageDto.setRecivedDate(msg.getMessageDate().toString());

			String writeExcel = writeExcelToPath(msg, tempExcelPath);

			if (writeExcel == null) {
				messageDto.setStatus("Done");
			} else {
				messageDto.setStatus("Error");
				messageDto.setErrMsg(writeExcel);
			}
			return messageDto;
		} catch (Exception e) {
			messageDto.setStatus("Error");
			messageDto.setErrMsg(e.getMessage());
			return messageDto;
		} finally {
			closeMessage(msg);
		}
	}

	private static String writeExcelToPath(MAPIMessage msg, String tempExcelPath) {
		AttachmentChunks[] attachments = msg.getAttachmentFiles();
		if (attachments.length > 0) {
			for (AttachmentChunks attachmentChunks : attachments) {
				StringChunk ext = attachmentChunks.getAttachExtension();
				System.out.println(ext.getValue());
				if (ext.getValue().equalsIgnoreCase(".xlsx") || ext.getValue().equalsIgnoreCase(".xls")) {
					boolean excelStatus = writeExcelFile(tempExcelPath, attachmentChunks);
					if (excelStatus)
						return null;
					else
						return "cann't write Excel file";
				}
			}
		}
		return "No attached";
	}

}
