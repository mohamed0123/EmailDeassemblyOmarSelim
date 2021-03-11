package com.zhopy.dto;

import java.io.Serializable;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
public class InputDto implements Serializable {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private String originalEmail;
	private String manufacturerPartNumberMpn;
	private String manufacturerName;
	private String customerInternalPartNumber;
	private String productDescription;
	private String comment;
	private String correctedManufacturerNameIfNecessarySeeComments;
	private String correctedMpnToBeFilledIfMpnIsIncorrectOrInvalid;
	private String lifecycleStatus;
	private String ltbDateTheLastDateByWhenTheCustomerCanOrderThePart;
	private String reasonForNrndDiscontinuedObsoletedIfPartIsNrndDiscontinuedObsoleted;
	private String partDesignType;
	private String rohs;
	private String euRohsExemptionListClickOnEmbeddedLink;
	private String rohs2015863NewAdded4PhthalatesStatusSelectOption1Yes2No;
	private String activeRohsReplacementMpn;
	private String formFitFunctionCompatibility;
	private String euRohsExemptionListForReplacementPartReferToLink;
	private String correctedManufacturerNameIfNecessary;
}
