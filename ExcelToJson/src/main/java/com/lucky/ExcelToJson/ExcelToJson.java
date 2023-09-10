package com.lucky.ExcelToJson;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
 
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
 
public class ExcelToJson {
 
    private ObjectMapper mapper = new ObjectMapper();
 
    public JsonNode excelToJson(File file) {
        // hold the excel data sheet wise
        ObjectNode excelData = mapper.createObjectNode();
        FileInputStream fis = null;
        Workbook workbook = null;
        try {
            // Creating file input stream
            fis = new FileInputStream(file);
 
            String filename = file.getName().toLowerCase();
            if (filename.endsWith(".xls") || filename.endsWith(".xlsx")) {
                // creating workbook object based on excel file format
                if (filename.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                } else {
                    workbook = new XSSFWorkbook(fis);
                }
                Sheet sheet = workbook.getSheet("Sheet1");
                String sheetName = sheet.getSheetName();
                List<String> headers = new ArrayList<String>();
                ArrayNode sheetData = mapper.createArrayNode();
                // Reading each row of the sheet
                for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                    Row row = sheet.getRow(j);
                    if (j == 0) {
                        // reading sheet header's name
                        for (int k = 0; k < row.getLastCellNum(); k++) {
                            headers.add(row.getCell(k).getStringCellValue());
                        }
                    } else {
                        // reading work sheet data
                        ObjectNode rowData = mapper.createObjectNode();
                        for (int k = 0; k < headers.size(); k++) {
                            Cell cell = row.getCell(k);
                            String headerName = headers.get(k);
                            if (cell != null) {
                            	rowData.put(headerName, cell.getStringCellValue());
                            } else {
                                rowData.put(headerName, "");
                            }
                        }
                        sheetData.add(rowData);
                    }
                }
                excelData.set(sheetName, sheetData);
                return excelData;
            } else {
                throw new IllegalArgumentException("File format not supported.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
 
        }
        return null;
    }
 
    public static void main(String[] args) {
 
        // Creating a file object with specific file path
        File file = new File("C:\\Users\\ADMIN\\Downloads\\ExcelToJson.xlsx");
        ExcelToJson converter = new ExcelToJson();
        JsonNode data = converter.excelToJson(file);
        System.out.println("Excel file contains the Data:\n" + data);
 
    }
}
