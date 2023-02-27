package com.example.payco.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.List;

@Slf4j
@Component
public class ExcelToModelConvertUtil {

    /*
        1. 엑셀 헤더 1개 있는 케이스
        2. 헤더 제외 후 2번째 Row부터 데이터 조회
     */
    public <T> List<T> excelToModelConvert(T t, final MultipartFile excelFile, int rowCount) throws Exception {

        List<T> excelToModelList = new ArrayList<>();

        try {
            OPCPackage opcPackage = OPCPackage.open(excelFile.getInputStream());
            XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);

            // 첫번째 시트 조회
            XSSFSheet sheet = workbook.getSheetAt(0);

            // 헤더 제외 후 2번째 Row부터 데이터 조회
            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum == 0) {
                return excelToModelList;
            }

            for (int i = rowCount; i <= lastRowNum; i++) {
                XSSFRow row = sheet.getRow(i);

                T tClass = (T) t.getClass().newInstance();

                int cellIndex = 0;
                for (Field field : t.getClass().getDeclaredFields()) {

                    field.setAccessible(true);

                    XSSFCell cell = row.getCell(cellIndex);

                    if (null != cell) {
                        if (field.getType().equals(String.class)) {
                            if (cell.getCellTypeEnum() != CellType.STRING) {
                                cell.setCellType(CellType.STRING);
                            }
                            field.set(tClass, cell.getStringCellValue());
                        } else if (field.getType().equals(Integer.class) || field.getType().equals(int.class)) {
                            field.set(tClass, (int) cell.getNumericCellValue());
                        } else if (field.getType().equals(Long.class) || field.getType().equals(long.class)) {
                            field.set(tClass, (long) cell.getNumericCellValue());
                        } else if (field.getType().equals(Double.class) || field.getType().equals(double.class)) {
                            field.set(tClass, cell.getNumericCellValue());
                        } else if (field.getType().equals(Float.class) || field.getType().equals(float.class)) {
                            field.set(tClass, cell.getNumericCellValue());
                        } else if (field.getType().equals(LocalDate.class)) {
                            field.set(tClass, cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate());
                        } else if (field.getType().equals(LocalDateTime.class)) {
                            field.set(tClass, cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
                        } else {
                            field.set(tClass, cell.getErrorCellString());
                        }
                    }

                    cellIndex++;

                }

                excelToModelList.add(tClass);
            }
        } catch (Exception ex) {
            throw new Exception("Excel 업로드 중에 오류가 발생했습니다. 오류 메시지: " + ex.getMessage() + " / class = " + t.getClass());
        }

        return excelToModelList;
    }
}
