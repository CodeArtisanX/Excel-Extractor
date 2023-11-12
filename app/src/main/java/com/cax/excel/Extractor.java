/*
 * Extractor.java
 *
 * Description:
 * This utility class provides methods for extracting field values from an Excel file based on specified rules.
 * It utilizes Apache POI for Excel handling and FastJSON for JSON serialization.
 *
 * Author:CodeArtisanX
 * Date: November 11, 2023
 */
package com.cax.excel;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Extractor utility class for extracting field values from an Excel file.
 */
public class Extractor {

    /**
     * Extracts field values from the specified Excel sheet based on the provided rules.
     *
     * @param sheetIndex The index of the Excel sheet.
     * @param excelFile  The Excel file from which to extract data.
     * @param rules      The rules specifying how to extract field values.
     * @return A map containing extracted field values.
     * @throws Exception If an error occurs during the extraction process.
     */
    public static Map<String, Object> extractFieldValuesMap(int sheetIndex, File excelFile, JSONArray rules) throws Exception {
        Map<String, Object> result = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(excelFile.getPath());
             Workbook workbook = WorkbookFactory.create(fis)) {
            // Assume data is in the first sheet
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            for (int i = 0; i < rules.size(); i++) {
                JSONObject rule = rules.getJSONObject(i);
                String fieldName = rule.getString("name");
                CellLocation targetCellLocation = findCell(sheet, fieldName, rule);
                if (targetCellLocation != null) {
                    Object fieldValue = getCellValue(sheet, targetCellLocation, rule);
                    result.put(fieldName, fieldValue);
                }
            }
        }
        return result;
    }

    /**
     * Extracts field values from the specified Excel sheet and returns the result as a JSON string.
     *
     * @param sheetIndex The index of the Excel sheet.
     * @param excelFile  The Excel file from which to extract data.
     * @param rules      The rules specifying how to extract field values.
     * @return A JSON string containing extracted field values.
     * @throws Exception If an error occurs during the extraction process.
     */
    public static String extractFieldValuesString(int sheetIndex, File excelFile, JSONArray rules) throws Exception {
        Map<String, Object> result = Extractor.extractFieldValuesMap(sheetIndex, excelFile, rules);
        return JSON.toJSONString(result);
    }

    /**
     * Finds the cell in the specified sheet that matches the given field name based on the provided rule.
     *
     * @param sheet     The Excel sheet to search for the cell.
     * @param fieldName The name of the field to find.
     * @param rule      The rule specifying the search direction.
     * @return The location of the cell as a CellLocation object, or null if not found.
     */
    private static CellLocation findCell(Sheet sheet, String fieldName, JSONObject rule) {
        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            if (row != null) {
                for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
                    Cell cell = row.getCell(colIndex);

                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();

                        // Check if the field name matches
                        if (fieldName.equals(cellValue)) {
                            int targetColIndex = colIndex;
                            int targetRowIndex = rowIndex;

                            // Adjust the target column index and row index based on the relative position
                            if (rule.getBoolean("right")) {
                                targetColIndex++;
                            } else if (rule.getBoolean("left")) {
                                targetColIndex--;
                            } else if (rule.getBoolean("down")) {
                                targetRowIndex++;
                            } else if (rule.getBoolean("up")) {
                                targetRowIndex--;
                            }

                            return new CellLocation(targetRowIndex, targetColIndex);
                        }
                    }
                }
            }
        }
        return null;
    }

    /**
     * Gets the value of the cell at the specified location in the given sheet based on the provided rule.
     *
     * @param sheet         The Excel sheet containing the cell.
     * @param cellLocation  The location of the target cell.
     * @param rule          The rule specifying whether the field value is a list.
     * @return The value of the cell, either as a single value or a list of values.
     */
    private static Object getCellValue(Sheet sheet, CellLocation cellLocation, JSONObject rule) {
        int targetRowIndex = cellLocation.getRowIndex();
        int targetColIndex = cellLocation.getColIndex();

        // Get the value of the target cell
        Row targetRow = sheet.getRow(targetRowIndex);

        if (targetRow != null) {
            Cell targetCell = targetRow.getCell(targetColIndex);

            if (targetCell != null) {
                if (rule.getBoolean("list")) {
                    // If the field value is a list, read multiple rows of data downward
                    List<String> valueList = new ArrayList<>();
                    for (int j = targetRowIndex + 1; j <= sheet.getLastRowNum(); j++) {
                        Row valueRow = sheet.getRow(j);
                        if (valueRow != null) {
                            Cell valueCell = valueRow.getCell(targetColIndex);
                            if (valueCell != null && valueCell.getCellType() == CellType.STRING) {
                                valueList.add(valueCell.getStringCellValue());
                            } else {
                                break;  // Break out of the loop if the next row is empty
                            }
                        } else {
                            break;  // Break out of the loop if the next row is empty
                        }
                    }
                    return valueList;
                } else {
                    // If the field value is not a list, directly read the value of the cell
                    return targetCell.getStringCellValue();
                }
            }
        }
        return null;
    }
}


