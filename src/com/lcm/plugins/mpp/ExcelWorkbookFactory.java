package com.lcm.plugins.mpp;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import net.sf.mpxj.ProjectFile;
import org.apache.xmlbeans.XmlException;

import java.util.List;
import java.util.Map;
import java.util.function.Function;

public class ExcelWorkbookFactory {

    public static XSSFWorkbook createWorkbookFromProjectFile(ProjectFile projectFile, SheetConfigDTO sheetConfig) throws XmlException {
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create Task sheet (mandatory)
        createSheet(workbook, projectFile.getTasks(), sheetConfig.getTaskHeaders(), sheetConfig.getTaskSheetName());

        // Create Resource sheet (optional)
        if (sheetConfig.getResourceHeaders() != null && sheetConfig.getResourceSheetName() != null) {
            createSheet(workbook, projectFile.getResources(), sheetConfig.getResourceHeaders(), sheetConfig.getResourceSheetName());
        }

        // Create Assignment sheet (optional)
        if (sheetConfig.getAssignmentHeaders() != null && sheetConfig.getAssignmentSheetName() != null) {
            createSheet(workbook, projectFile.getResourceAssignments(), sheetConfig.getAssignmentHeaders(), sheetConfig.getAssignmentSheetName());
        }

        return workbook;
    }

    private static <T> void createSheet(XSSFWorkbook workbook, List<T> items, Map<String, Function<T, String>> headers, String sheetName) {
        XSSFSheet sheet = workbook.createSheet(sheetName);

        // Create header row
        XSSFRow headerRow = sheet.createRow(0);
        int cellIndex = 0;
        for (String header : headers.keySet()) {
            headerRow.createCell(cellIndex++).setCellValue(header);
        }

        // Create data rows
        int rowIndex = 1;
        for (T item : items) {
            XSSFRow row = sheet.createRow(rowIndex++);
            cellIndex = 0;
            for (Map.Entry<String, Function<T, String>> entry : headers.entrySet()) {
                row.createCell(cellIndex++).setCellValue(entry.getValue().apply(item));
            }
        }
    }
}
