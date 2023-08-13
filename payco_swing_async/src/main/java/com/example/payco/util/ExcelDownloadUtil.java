package com.example.payco.util;

import com.example.payco.model.ExcelDownloadResultDTO;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

@Slf4j
public class ExcelDownloadUtil {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private ExcelDownloadResultDTO resultDTO;

    public ExcelDownloadUtil(ExcelDownloadResultDTO resultDTO) {
        this.resultDTO = resultDTO;
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet();
    }

    public void export(HttpServletResponse response) throws Exception {
        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=" + resultDTO.getFileName();
        response.setHeader(headerKey, headerValue);

        setExcelContent(response);
    }

    private void setExcelContent(HttpServletResponse response) throws Exception {
        writeHeaderRow();
        writeDataRows();

        if ("".equals(resultDTO.getPassword()) && resultDTO.getPassword() == null) {
            try (ServletOutputStream outputStream = response.getOutputStream()) {
                workbook.write(outputStream);
            }
        } else {
            protectedPassword(response);
        }

        workbook.close();
    }

    private void protectedPassword(HttpServletResponse response) throws Exception {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        workbook.write(byteArrayOutputStream);

        InputStream inputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
        POIFSFileSystem fs = new POIFSFileSystem();
        EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
        Encryptor encryptor = info.getEncryptor();

        encryptor.confirmPassword(resultDTO.getPassword());

        OPCPackage opc = OPCPackage.open(inputStream);
        OutputStream os = encryptor.getDataStream(fs);
        opc.save(os);
        opc.close();

        OutputStream outputStream = response.getOutputStream();
        fs.writeFilesystem(outputStream);
        outputStream.close();
        byteArrayOutputStream.close();
    }

    private void writeHeaderRow() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        List<String> headerList = resultDTO.getHeaderList();
        for (int i = 0; i < headerList.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headerList.get(i));
            cell.setCellStyle(style);
        }
    }

    private void writeDataRows() {
        List<List<String>> dataList = resultDTO.getDataList();
        for (int i = 0; i < dataList.size(); i++) {
            Row row = sheet.createRow(i + 1);

            List<String> rowDataList = dataList.get(i);
            for (int j = 0; j < rowDataList.size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(rowDataList.get(j));
            }
        }
    }
}
