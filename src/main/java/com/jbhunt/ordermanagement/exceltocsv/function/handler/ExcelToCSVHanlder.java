package com.jbhunt.ordermanagement.exceltocsv.function.handler;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

public class ExcelToCSVHanlder {

    public void handleFile(byte[] sourceBytes, File destFile) {
        if (sourceBytes == null || destFile == null) {
//            throw new JBHValidationException("Invalid input file or the destination file name");
        }

        StringBuilder fileContent = this.readFileContent(sourceBytes);
        if (fileContent != null && fileContent.length() > 0) {
            saveFile(fileContent, destFile);
//            deleteSourceFile(sourceFile);
        }

    }

    public void saveFile(StringBuilder fileContent, File destFileName) {
        try (PrintStream out = new PrintStream(new FileOutputStream(destFileName),
                true, ("UTF-8"))) {
            byte[] bom = {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
            out.write(bom);
            out.write(fileContent.toString().getBytes());
        } catch (Exception exception) {
//            throw new JBHException("Exception in writing file", exception);
        }
    }

    public StringBuilder readFileContent(byte[] bytes) {
        StringBuilder fileContent = new StringBuilder();
        try (InputStream inputStream = new ByteArrayInputStream(bytes);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            int numberOfWb = workbook.getNumberOfSheets();
            for (int sheetIndex = 0; sheetIndex < numberOfWb; sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                fileContent.append(iterateRows(sheet));
            }

        } catch (IOException ioException) {
//            throw new JBHException(ioException.getMessage(), ioException);
        } catch (Exception exception) {
//            log.error("Exception in reading file", exception);
//            throw new JBHException(exception.getMessage(), exception);
        }
        return fileContent;
    }

    private StringBuilder iterateRows(Sheet sheet) {
        int totalRows = sheet.getLastRowNum();
        StringBuilder sheetContent = new StringBuilder();
        for (int rowIndex = 0; rowIndex <= totalRows; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                sheetContent.append(",");
                continue;
            }
            sheetContent.append(iterateCells(row));
            sheetContent.append("\n");
        }
        return sheetContent;
    }

    private StringBuilder iterateCells(Row row) {
        DataFormatter formatter = new DataFormatter();
        int totalCells = row.getLastCellNum();
        boolean firstCell = true;
        StringBuilder rowContent = new StringBuilder();
        for (int cellIndex = 0; cellIndex < totalCells; cellIndex++) {
            Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

            if (cell == null) {
                rowContent.append(",");
                continue;
            }
            if (!firstCell) {
                rowContent.append(",");
            }
            String cellValue = formatter.formatCellValue(cell);
            rowContent.append(cellValue);
            firstCell = false;
        }
        return rowContent;
    }

}
