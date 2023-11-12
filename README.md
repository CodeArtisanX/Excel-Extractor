# Excel-Extractor
This project demonstrates rule-based extraction of field values from an Excel file.

**Other**： [中文版](README_CN.md) .
## Quick Start

### Maven Dependency

To use this project in your Maven-based application, you can add the following dependency to your `pom.xml`:

```xml
<dependency>
  <groupId>com.cax</groupId>
  <artifactId>excel-extractor</artifactId>
  <version>1.0.0</version>
</dependency>
```
## Quick Usage Example
The MainTest class provides a simple example of how to use the Extractor utility to extract field values from an Excel file based on specified rules. The RuleBuilder class is utilized to create extraction rules for testing purposes.
### Usage
1. **Input Excel File**: Provide the path to the Excel file containing the data to be processed. Update the `excelFile` variable in the `main` method with the correct file path.

    ```java
    // Assume there is an existing Excel file with data
    File excelFile = new File("/path/to/your/excel/file.xlsx");
    ```

2. **Create Extraction Rules**: Use the `RuleBuilder` to create rules for each field to be extracted. The `right()`, `left()`, `up()`, `down()`, and `list()` methods are available to specify the relative positions and whether the field is part of a list.

    ```java
    // Create rules using the RuleBuilder
    JSONArray rules = new JSONArray();
    rules.add(new RuleBuilder("ColumnName").right().down().list().build());
    rules.add(new RuleBuilder("ColumnName1").right().build());
    ```

3. **Run the Test**: Execute the `main` method to run the test. The extracted data will be printed in the console.

    ```java
    // Read Excel file and extract field values based on rules
    String result = Extractor.extractFieldValuesString(0, excelFile, rules);
    System.out.println(result);
    ```

## Note

- This test assumes the existence of an Excel file with the specified path. Update the file path as needed.

- Ensure that the `Extractor` utility and `RuleBuilder` class are correctly implemented and available in your project.

- This is a basic example, and you may need to adapt the code based on the actual structure and content of your Excel file.

## License

This project is licensed under the [Apache-2.0 license] - see the [LICENSE.md](LICENSE.md) file for details.
