package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {

    public static void main(String[] args) {


        String filePath = "C:\\Users\\User\\Downloads\\Test.xls";
        int columnIndex = 3;
        String tempValue = null;

        try {
            FileInputStream file = new FileInputStream(filePath);
            Workbook workbook = new HSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);
            CellStyle style = workbook.createCellStyle();

            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    if (cell.getStringCellValue().contains("Total")) {
                        continue;
                    }
                    tempValue = cell.getStringCellValue();
                    style = cell.getCellStyle();
                }
                Cell cellForCheck = row.getCell(7);
                if (cell != null && cellForCheck.getCellType() == CellType.BLANK) {
                    continue;
                } else {
                    cell = row.createCell(columnIndex);
                    cell.setCellValue(tempValue);
                    cell.setCellStyle(style);

                    for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
                        CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                        if (mergedRegion.isInRange(cell.getRowIndex(), columnIndex)) {
                            sheet.removeMergedRegion(i);
                            break;
                        }
                    }
                }

                for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
                    CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                    if (mergedRegion.isInRange(cell.getRowIndex(), 7)) {
                        sheet.removeMergedRegion(i);
                        break;
                    }
                }
            }

            sheet.autoSizeColumn(3);
            sheet.autoSizeColumn(7);
            sheet.setColumnWidth(0, 10000);
            sheet.setColumnWidth(8, 0);
            sheet.setColumnWidth(9, 0);
            sheet.setColumnWidth(10, 0);
            sheet.setColumnWidth(1, 0);
            sheet.setColumnWidth(4, 0);

            FileOutputStream fileOut = new FileOutputStream(filePath);
            workbook.write(fileOut);
            fileOut.close();
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("The excel file has been changed successfully");
    }

}