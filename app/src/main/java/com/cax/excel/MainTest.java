/*
 * MainTest.java
 *
 * Description:
 * This is a test class for demonstrating the usage of the Extractor utility to extract field values
 * from an Excel file based on specified rules. It includes sample rules for extracting data.
 *
 * Author: CodeArtisanX
 * Date: November 11, 2023
 */
package com.cax.excel;

import com.alibaba.fastjson.JSONArray;

import java.io.File;

/**
 * MainTest class for testing the Extractor utility.
 */
public class MainTest {

    public static void main(String[] args) {
        try {
            // Assume there is an existing Excel file with data
            File excelFile = new File("/path/to/your/excel/file.xlsx");

            // Create rules using the RuleBuilder
            JSONArray rules = new JSONArray();
            rules.add(new RuleBuilder("ColumnName").right().down().list().build());
            rules.add(new RuleBuilder("ColumnName1").right().build());

            // Read Excel file and extract field values based on rules
            String result = Extractor.extractFieldValuesString(0, excelFile, rules);

            System.out.println(result);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
