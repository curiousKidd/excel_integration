package com.example.payco.model;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

@Getter
@Setter
@ToString
public class BayDTO {

    private String empty;
    private String seq;
    private LocalDate useDate;
    private String name;
    private String usePlace;
    private String ticketType;
    private int ticketPaymentAmount;
    private String names;

    public String getUseDate() {
        DateTimeFormatter DATE_FORMAT_YYYY_MM_DD = DateTimeFormatter.ofPattern("yyyy.MM.dd", Locale.KOREA);
        return this.useDate != null ? this.useDate.format(DATE_FORMAT_YYYY_MM_DD) : "";
    }
}
