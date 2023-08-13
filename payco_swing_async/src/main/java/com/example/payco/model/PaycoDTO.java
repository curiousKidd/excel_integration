package com.example.payco.model;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
public class PaycoDTO {

    private String tranNumber;
    private String tranDate;
    private String companyNumber;
    private String name;
    private String group;
    private String rank;
    private String usePlace;
    private String paymentType;
    private String tranType;
    private int totalPaymentAmount;
    private int ticketPaymentAmount;
    private String ticketType;
    
}
