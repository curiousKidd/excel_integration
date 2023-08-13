package com.example.payco.model;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
@Builder
public class ExcelDTO {

    private String tranNumber;
    private String tranDate;
    private String companyNumber;
    private String name;
    private String usePlace;
    private String paymentType;
    private String tranType;
    private int totalPaymentAmount;
    private int ticketPaymentAmount;
    private String ticketType;
    private String names;

}
