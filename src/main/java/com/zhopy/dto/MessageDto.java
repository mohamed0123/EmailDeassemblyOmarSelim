package com.zhopy.dto;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class MessageDto {

	private String cc;
	private String from;
	private String to;
	private String recivedDate;
	private String status;
	private String errMsg;
}
