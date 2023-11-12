# Excel-Extractor

该项目演示了基于规则的从Excel文件中提取字段值的方法。

**其他**: [英文版](README.md).

## 快速开始

### Maven 依赖

要在基于Maven的应用程序中使用该项目，您可以将以下依赖项添加到您的 `pom.xml` 文件中：

```xml
<dependency>
  <groupId>com.cax</groupId>
  <artifactId>excel-extractor</artifactId>
  <version>1.0.0</version>
</dependency>
```

## 快速使用示例

MainTest类提供了如何使用Extractor实用程序根据指定规则从Excel文件中提取字段值的简单示例。RuleBuilder类用于为测试目的创建提取规则。

### 使用步骤

1. **输入Excel文件**：提供包含要处理的数据的Excel文件的路径。在`main`方法中更新`excelFile`变量，以正确设置文件路径。

    ```java
    // 假设存在包含数据的Excel文件
    File excelFile = new File("/path/to/your/excel/file.xlsx");
    ```

2. **创建提取规则**：使用`RuleBuilder`为要提取的每个字段创建规则。可使用`right()`、`left()`、`up()`、`down()`和`list()`方法指定相对位置以及字段是否属于列表。

    ```java
    // 使用RuleBuilder创建规则
    JSONArray rules = new JSONArray();
    rules.add(new RuleBuilder("ColumnName").right().down().list().build());
    rules.add(new RuleBuilder("ColumnName1").right().build());
    ```

3. **运行测试**：执行`main`方法运行测试。提取的数据将在控制台中打印出来。

    ```java
    // 读取Excel文件并根据规则提取字段值
    String result = Extractor.extractFieldValuesString(0, excelFile, rules);
    System.out.println(result);
    ```

## 注意事项

- 此测试假设存在具有指定路径的Excel文件。根据需要更新文件路径。

- 确保`Extractor`实用程序和`RuleBuilder`类在您的项目中正确实现并可用。

- 这是一个基本示例，您可能需要根据实际的Excel文件结构和内容调整代码。

## 许可证

本项目基于[Apache-2.0许可证]授权 - 有关详细信息，请参见[LICENSE](LICENSE)文件。
