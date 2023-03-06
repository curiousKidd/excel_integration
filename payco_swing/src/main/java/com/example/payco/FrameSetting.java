package com.example.payco;

import com.example.payco.model.BayDTO;
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
import java.util.stream.Collectors;

public class FrameSetting extends JFrame implements ActionListener {

    private ExcelToModelConvertUtil excelToModelConvertUtil = new ExcelToModelConvertUtil();

    private JFileChooser fileComponent = new JFileChooser("C:\\Users\\ywjeon.ITEMBAY\\Desktop\\식대");

    private JButton paycoButton = new JButton("payco Excel");
    private JButton bayButton = new JButton("bay Excel");
    private JButton btnSaveButton = new JButton("저장");

    private JTextField paycoLabel = new JTextField(40);
    private JTextArea bayLabel = new JTextArea();

    private List<File> bayFile = new ArrayList<>();

    private MultipartFile paycoMultipartFile = null;
    private List<MultipartFile> bayMultipartFile = new ArrayList<>();

    JScrollPane scrollPane = new JScrollPane(bayLabel);

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

        paycoLabel.setEditable(false);
        bayLabel.setEditable(false);

        getContentPane().setLayout(null);

        paycoButton.setBounds(62, 30, 120, 20);
        getContentPane().add(paycoButton);

        paycoLabel.setBounds(190, 30, 310, 20);
        getContentPane().add(paycoLabel);
        paycoLabel.setColumns(10);

        bayButton.setBounds(62, 70, 120, 20);
        getContentPane().add(bayButton);

        scrollPane.setBounds(62, 100, 440, 93);
        getContentPane().add(scrollPane);

        btnSaveButton.setBounds(420, 200, 80, 20);
        getContentPane().add(btnSaveButton);
    }

    public void start() {
        paycoButton.addActionListener(this);
        bayButton.addActionListener(this);
        btnSaveButton.addActionListener(this);
        fileComponent.setFileFilter(new FileNameExtensionFilter("xlsx", "xlsx", "xls")); // 확장자 .xlsx, xls만 선택가

    }

    public void actionPerformed(ActionEvent arg0) {
        try {
            if (arg0.getSource() == paycoButton) {
                fileComponent.setMultiSelectionEnabled(false); // 다중 선택 불가 설정
                if (fileComponent.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                    paycoMultipartFile = this.getMultipartFile(fileComponent.getSelectedFile());
                    paycoLabel.setText(fileComponent.getSelectedFile().getPath());
                }
            } else if (arg0.getSource() == bayButton) {
                fileComponent.setMultiSelectionEnabled(true);
                if (fileComponent.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                    File[] selectedFiles = fileComponent.getSelectedFiles();

                    for (File file : bayFile) {
                        if (Arrays.stream(selectedFiles).anyMatch(f -> f.getPath().equals(file.getPath()))) {
                            JOptionPane.showMessageDialog(null, "파일 중복");
                            return;
                        }
                    }

                    bayFile.addAll(Arrays.asList(selectedFiles));

                    Arrays.stream(selectedFiles).forEach(f -> {
                        try {
                            bayLabel.append(f.getPath() + "\n");
                            bayMultipartFile.add(this.getMultipartFile(f));
                        } catch (IOException e) {
                            JOptionPane.showMessageDialog(null, "파일 생성중 오류가 발생했습니다.");
                            JOptionPane.showMessageDialog(null, e.toString());
                        }
                    });

                    bayLabel.setCaretPosition(bayLabel.getDocument().getLength());  // 맨아래로 스크롤한다.
                }
            } else if (arg0.getSource() == btnSaveButton) {
                this.excelData(paycoMultipartFile, bayMultipartFile);
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
            @RequestParam(value = "fileUploadPayco") MultipartFile paycoExcelFile,
            @RequestParam(value = "fileUploadBay") List<MultipartFile> bayExcelFile) throws Exception {

        java.util.List<PaycoDTO> paycoExcelData = excelToModelConvertUtil.excelToModelConvert(new PaycoDTO(), paycoExcelFile, 1);

        List<java.util.List<BayDTO>> bayExcelData = new ArrayList<>();

        for (MultipartFile multipartFile : bayExcelFile) {

            bayExcelData.add(excelToModelConvertUtil.excelToModelConvert(new BayDTO(), multipartFile, 5));
        }
        java.util.List<ExcelDTO> excelSum = excelSum(paycoExcelData, bayExcelData);

        exportExcel(excelSum);

    }

    private java.util.List<ExcelDTO> excelSum(java.util.List<PaycoDTO> paycoDTOS, List<java.util.List<BayDTO>> bayDTOS) {

        Comparator<PaycoDTO> compare = Comparator
                .comparing(PaycoDTO::getTicketType)
                .thenComparing(PaycoDTO::getTranNumber);

        Set<String> set = new HashSet<>();

        paycoDTOS.stream()
                .sorted(Comparator.comparing(PaycoDTO::getTranType))
                .filter(f -> f.getTranType().equals("취소"))
                .forEach(fe -> set.add(fe.getTranNumber()));

        java.util.List<ExcelDTO> excelDTOS = new ArrayList<>();

        for (List<BayDTO> bayDTO : bayDTOS) {

            excelDTOS = paycoDTOS.stream()
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
                            .names(bayDTO.stream()
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

        return excelDTOS;
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
