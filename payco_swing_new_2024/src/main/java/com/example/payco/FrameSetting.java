package com.example.payco;

import com.example.payco.model.ExcelDTO;
import com.example.payco.model.PaycoDTO;
import com.example.payco.util.ExcelToModelConvertUtil;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItem;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.nio.file.Files;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Collectors;

public class FrameSetting extends JFrame implements ActionListener {

    private ExcelToModelConvertUtil excelToModelConvertUtil = new ExcelToModelConvertUtil();

    private JFileChooser fileComponent = new JFileChooser("C:\\Users\\ywjeon.ITEMBAY\\Desktop\\식대");

    private JButton paycoButton = new JButton("payco Excel");
    private JButton btnSaveButton = new JButton("저장");

    private JLabel label = new JLabel("Price");

    private JTextField paycoLabel = new JTextField(40);
    private JTextField priceLabel = new JTextField(40);

    private MultipartFile paycoMultipartFile = null;

    private static List<PaycoDTO> paycoExcelData = new ArrayList<>();
    private static final DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy.MM.dd HH:mm:ss");

    private static int price = 0;

    public FrameSetting() {
        this.init();
        this.start();
        this.setSize(600, 300);
        this.setLocationRelativeTo(null);
        this.setVisible(true);

    }

    public void init() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setBounds(100, 100, 554, 239);
        setTitle("페이코 합산하는거에 시간 들이지 마요");

        getContentPane().setLayout(null);

        // 페이코 버튼
        paycoButton.setBounds(62, 30, 120, 20);
        label.setHorizontalAlignment(JLabel.CENTER);
        getContentPane().add(paycoButton);

        // 페이코 라벨
        paycoLabel.setBounds(190, 30, 310, 20);
        paycoLabel.setEditable(false);      // 수정 불가능하게
        getContentPane().add(paycoLabel);
        paycoLabel.setColumns(10);

        label.setBounds(62, 60, 120, 20);
        getContentPane().add(label);

        priceLabel.setBounds(190, 60, 310, 20);
        priceLabel.setText("9000");
        getContentPane().add(priceLabel);
        paycoLabel.setColumns(10);

        btnSaveButton.setBounds(420, 200, 80, 20);
        getContentPane().add(btnSaveButton);
    }

    public void start() {
        paycoButton.addActionListener(this);
        btnSaveButton.addActionListener(this);
        fileComponent.setFileFilter(new FileNameExtensionFilter("xlsx", "xlsx", "xls")); // 확장자 .xlsx, xls만 선택가
        fileComponent.setMultiSelectionEnabled(false); // 다중 선택 불가 설정
    }

    public void actionPerformed(ActionEvent arg0) {
        try {
            if (arg0.getSource() == paycoButton) {
                if (fileComponent.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                    paycoMultipartFile = this.getMultipartFile(fileComponent.getSelectedFile());
                    paycoLabel.setText(fileComponent.getSelectedFile().getPath());
                }
            } else if (arg0.getSource() == btnSaveButton) {
                this.excelData(paycoMultipartFile);
                JOptionPane.showMessageDialog(null, "파일 생성이 완료되었습니다.");

            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "파일 생성중 오류가 발생했습니다.");
            JOptionPane.showMessageDialog(null, e.toString());
        }
    }

    private MultipartFile getMultipartFile(File file) throws IOException {
        FileItem fileItem = new DiskFileItem("originFile", Files.probeContentType(file.toPath()), false, file.getName(), (int) file.length(), file.getParentFile());

        InputStream input = new FileInputStream(file);
        OutputStream os = fileItem.getOutputStream();
        IOUtils.copy(input, os);

        MultipartFile mFile = new CommonsMultipartFile(fileItem);

        return mFile;
    }

    public void excelData(
            @RequestParam(value = "fileUploadPayco") MultipartFile paycoExcelFile) throws Exception {

        paycoExcelData = excelToModelConvertUtil.excelToModelConvert(new PaycoDTO(), paycoExcelFile, 1);

        List<ExcelDTO> excelSum = getExcelData();

        exportExcel(excelSum);

    }

    // 엑셀 데이터 생성
    private List<ExcelDTO> getExcelData() {

        price = Integer.parseInt(priceLabel.getText());

        java.util.List<ExcelDTO> excelDTOS = new ArrayList<>();
        String excelNames = "";

        Comparator<PaycoDTO> compare = Comparator
                .comparing(PaycoDTO::getTicketType)
                .thenComparing(PaycoDTO::getTranNumber);

        Set<String> set = new HashSet<>();

        paycoExcelData.stream()
                .sorted(Comparator.comparing(PaycoDTO::getTranType))
                .filter(f -> "취소".equals(f.getTranType()))
                .forEach(fe -> set.add(fe.getTranNumber()));

        excelDTOS = paycoExcelData.stream()
                .sorted(compare)
                .filter(f -> set.add(f.getTranNumber()) && "승인".equals(f.getTranType()))
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
                        //                        .names(m.getTicketPaymentAmount() > price ? getNames(m, price, paycoDTOS) : "")
                        .names(getNames(m))
                        .build()
                )
                .collect(Collectors.toList());

        return excelDTOS;
    }

    // 통합 사용자 이름 가져오기
    // paycoDTO : 승인 데이터
    private String getNames(PaycoDTO paycoDTO) {
        HashMap<String, Integer> map = new HashMap<>();
        StringBuilder sb = new StringBuilder();

        paycoExcelData.stream()
                .filter(f ->
                        f.getName().equals(paycoDTO.getName())
                                && f.getTranDate().substring(0, 10).equals(paycoDTO.getTranDate().substring(0, 10))
                                && f.getTicketType().equals(paycoDTO.getTicketType())
                                && !"취소".equals(f.getTranType())
                )
                .map(fm -> {
                    if (!"승인".equals(fm.getTranType())) {
                        if ("받기".equals(fm.getTranType())) {
                            map.putAll(amountTracking(fm));
                            //                            map.put(fm.getUsePlace(), map.getOrDefault(fm.getUsePlace(), 0) + fm.getTicketPaymentAmount());
                        } else {
                            map.put(fm.getUsePlace(), map.getOrDefault(fm.getUsePlace(), 0) - fm.getTicketPaymentAmount());
                        }
                    } else if (paycoDTO.getTranNumber().equals(fm.getTranNumber())) {
                        map.put(fm.getName(), map.getOrDefault(fm.getName(), 0) + fm.getTicketPaymentAmount());

                    }
                    return fm;
                })
                .collect(Collectors.toList());

        Map<String, Integer> filteredMap = map.entrySet()
                .stream()
                .filter(f -> f.getValue() > 0)
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));

        if (filteredMap.size() > 1) {
            for (String key : filteredMap.keySet()) {
                sb.append(key);
                sb.append(", ");
            }
        }

        return sb.toString();
    }

    // 받기 금액 추적
    // PaycoDTO dto : 받기 내역 DTO
    private Map<String, Integer> amountTracking(PaycoDTO dto) {
        HashMap<String, Integer> map = new HashMap<>();
        StringBuilder sb = new StringBuilder();
        AtomicBoolean approvalCheck = new AtomicBoolean(false);

        Comparator<PaycoDTO> compare = Comparator
                .comparing(PaycoDTO::getTicketType)
                .thenComparing(PaycoDTO::getTranNumber).reversed();

        map.put(dto.getUsePlace(), dto.getTicketPaymentAmount());

        paycoExcelData.stream()
                .sorted(compare)
                .filter(f ->
                        f.getName().equals(dto.getUsePlace())
                                && f.getTranDate().substring(0, 10).equals(dto.getTranDate().substring(0, 10))
                                && LocalDateTime.parse(f.getTranDate(), DATE_TIME_FORMATTER).compareTo(LocalDateTime.parse(dto.getTranDate(), DATE_TIME_FORMATTER)) < 0
                                && f.getTicketType().equals(dto.getTicketType())
                                && !"취소".equals(f.getTranType())
                )
                .map(m -> {
                    if (!approvalCheck.get()) {
                        if ("받기".equals(m.getTranType())) {
                            map.putAll(amountTracking(m));
                            //                        map.put(m.getUsePlace(), map.getOrDefault(m.getUsePlace(), 0) + m.getTicketPaymentAmount());
                        } else if ("보내기".equals(m.getTranType())) {
                            map.put(m.getUsePlace(), map.getOrDefault(m.getUsePlace(), 0) - m.getTicketPaymentAmount());
                        } else if ("승인".equals(m.getTranType()) && m.getTicketPaymentAmount() % price == 0) {
                            approvalCheck.set(true);
                        }
                    }
                    return m;
                })
                .collect(Collectors.toList());

        //Map filter
        Map<String, Integer> filteredMap = map.entrySet()
                .stream()
                .filter(f -> f.getValue() > 0)
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));

        return filteredMap;
    }

    // 엑셀 생성
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

        //헤더 생성
        curRow = sh.createRow(0);
        curRow.createCell(0).setCellValue("거래번호");
        curRow.createCell(1).setCellValue("거래일시");
        curRow.createCell(2).setCellValue("사번");
        curRow.createCell(3).setCellValue("이름");
        curRow.createCell(4).setCellValue("사용처");
        curRow.createCell(5).setCellValue("결제방식");
        curRow.createCell(6).setCellValue("거래타입");
        curRow.createCell(7).setCellValue("총 결제금액");
        curRow.createCell(8).setCellValue("식권 결제금액");
        curRow.createCell(9).setCellValue("식권명");
        curRow.createCell(10).setCellValue("복합사용자");

        for (int i = 0; i < list.size(); i++) {
            curRow = sh.createRow(i + 1);    // row 생성
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

            String strFilePath = "C:\\PAYCO\\";

            folderCreateTest(strFilePath);

            DateTimeFormatter DATE_FORMAT_YYYY_MM_DD_HH_MM_SS = DateTimeFormatter.ofPattern("yyyy.MM.dd.HH.mm.ss", Locale.KOREA);
            String fileName = LocalDateTime.now().format(DATE_FORMAT_YYYY_MM_DD_HH_MM_SS) + "_payco.xlsx";

            FileOutputStream fOut = new FileOutputStream(strFilePath + fileName);
            wb.write(fOut);

        } catch (IOException e) {
            e.printStackTrace();
            System.exit(0);
        }

    }

    // 폴더 생성
    private void folderCreateTest(String folderPath) {

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
