package com.example.payco.model;

import lombok.*;

import java.util.List;

@Setter
@Getter
@Builder
@ToString
@NoArgsConstructor
@AllArgsConstructor
public class ExcelDownloadResultDTO {
    private String fileName;
    private String password;
    private String sheetName;
    private List<String> headerList;
    private List<List<String>> dataList;

}
