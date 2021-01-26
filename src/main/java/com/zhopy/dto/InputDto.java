package com.zhopy.dto;

import java.io.Serializable;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

@AllArgsConstructor
@Getter
@Setter
public class InputDto implements Serializable {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private String partNumber;
	private String description;
	private String lcStatus;
	private String originalEmail;
}
