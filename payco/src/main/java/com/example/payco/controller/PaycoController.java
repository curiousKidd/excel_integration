package com.example.payco.controller;

import com.example.payco.model.BayDTO;
import com.example.payco.model.ExcelDTO;
import com.example.payco.model.PaycoDTO;
import com.example.payco.util.ExcelToModelConvertUtil;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.YearMonth;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

@Controller
@RequiredArgsConstructor
public class PaycoController {

    private final ExcelToModelConvertUtil excelToModelConvertUtil;

    @GetMapping("/")
    public String main() {
        return "excelInput";
    }

    @PostMapping("/excel")
    public String excelData(
            @RequestParam(value = "fileUploadPayco") MultipartFile paycoExcelFile,
            @RequestParam(value = "fileUploadBay") MultipartFile bayExcelFile,
            HttpServletResponse response,
            Model model) throws Exception {

        List<PaycoDTO> paycoExcelData = excelToModelConvertUtil.excelToModelConvert(new PaycoDTO(), paycoExcelFile, 1);
        List<BayDTO> bayExcelData = excelToModelConvertUtil.excelToModelConvert(new BayDTO(), bayExcelFile, 5);
        List<ExcelDTO> excelSum = excelSum(paycoExcelData, bayExcelData);

        //        exportExcel(excelSum);

        //alert
        //        response.setContentType("text/html; charset=UTF-8");
        //        PrintWriter out = response.getWriter();
        //        out.println("<script>alert('파일 생성 완료');</script>");
        //        out.flush();
        //        response.flushBuffer();
        //        out.close();

        model.addAttribute("paycoExcelData", paycoExcelData);
        model.addAttribute("bayExcelData", bayExcelData);
        model.addAttribute("excelSum", excelSum);

        return "excelData";

    }

    private List<ExcelDTO> excelSum(List<PaycoDTO> paycoDTOS, List<BayDTO> bayDTOS) {

        Comparator<PaycoDTO> compare = Comparator
                .comparing(PaycoDTO::getTicketType)
                .thenComparing(PaycoDTO::getTranNumber);

        Set<String> set = new HashSet<>();

        paycoDTOS.stream()
                .sorted(Comparator.comparing(PaycoDTO::getTranType))
                .filter(f -> f.getTranType().equals("취소"))
                .forEach(fe -> set.add(fe.getTranNumber()));

        return paycoDTOS.stream()
                .sorted(compare)
                .filter(f -> set.add(f.getTranNumber()))
                .map(m -> ExcelDTO.builder()
                        .tranNumber(m.getTranNumber())
                        .tranDate(m.getTranDate())
                        .companyNumber(m.getCompanyNumber())
                        .name(m.getName())
                        .usePlace(m.getUsePlace())
                        .paymentType(m.getPaymentType())
                        .tranType(m.getTranType())
                        .totalPaymentAmount(m.getTotalPaymentAmount())
                        .ticketPaymentAmount(m.getTicketPaymentAmount())
                        .ticketType(m.getTicketType())
                        .names(bayDTOS.stream()
                                .filter(f ->
                                        f.getName().equals(m.getName())
                                                && f.getTicketPaymentAmount() == m.getTicketPaymentAmount()
                                                && f.getUseDate().equals(m.getTranDate().substring(0, 10))
                                                && f.getTicketType().equals(m.getTicketType())
                                )
                                .findFirst().orElseGet(BayDTO::new).getNames()
                        )
                        .build()
                ).collect(Collectors.toList());
    }

    private void exportExcel(List<ExcelDTO> list) {

        // Workbook 생성
        @SuppressWarnings("resource")
        XSSFWorkbook wb = new XSSFWorkbook(); // Excel 2007 이상

        XSSFSheet sh = wb.createSheet("Sheet1");

        XSSFRow row = null;
        XSSFCell cell = null;

        // Cell의 스타일을 적용한다.
        CellStyle csBase = wb.createCellStyle();

        // 글꼴
        XSSFFont fBase = wb.createFont();
        fBase.setFontName("나눔고딕코딩");       // Font
        fBase.setFontHeightInPoints((short) 8);  // Font
        csBase.setFont(fBase);

        XSSFRow curRow;

        for (int i = 0; i < list.size(); i++) {

            curRow = sh.createRow(i);    // row 생성
            curRow.createCell(0).setCellValue(list.get(i).getTranNumber());    // row에 각 cell 저장
            curRow.createCell(1).setCellValue(list.get(i).getTranDate());
            curRow.createCell(2).setCellValue(list.get(i).getCompanyNumber());
            curRow.createCell(3).setCellValue(list.get(i).getName());
            curRow.createCell(4).setCellValue(list.get(i).getUsePlace());
            curRow.createCell(5).setCellValue(list.get(i).getPaymentType());
            curRow.createCell(6).setCellValue(list.get(i).getTranType());
            curRow.createCell(7).setCellValue(list.get(i).getTotalPaymentAmount());
            curRow.createCell(8).setCellValue(list.get(i).getTicketPaymentAmount());
            curRow.createCell(9).setCellValue(list.get(i).getTicketType());
            curRow.createCell(10).setCellValue(list.get(i).getNames());
        }

        try {

            String strFilePath = "D:/excel/";

            folderCreateTest(strFilePath);

            String fileName = YearMonth.now().toString() + "_payco.xlsx";

            FileOutputStream fOut = new FileOutputStream(strFilePath + fileName);
            wb.write(fOut);

        } catch (IOException e) {
            e.printStackTrace();
            System.exit(0);
        }

    }

    private void folderCreateTest(String folderPath) {
        //        String folderPath = "D:\\Dev\\test\\test2222";

        File file = new File(folderPath);

        if (!file.exists()) {
            try {
                if (!file.isFile()) {
                    file.mkdirs(); //폴더 생성합니다.
                }
            } catch (Exception e) {
                System.out.println(e);
            }
        }
    }

}
