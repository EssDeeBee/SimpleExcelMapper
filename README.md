# SimpleExcelMapper

With this library, you can create an Excel file from data classes and convert an Excel report back into those classes.

# Data classes to Excel example:

```java
void createSalesReport() {
    ExcelMapper excelMapper = new ExcelMapperImpl();

    var fileName = "sales-out-" + LocalTime.now().format(DateTimeFormatter.ofPattern("hh.mm.ss")) + ".xlsx";
    List<Sales> sales = List.of(
            new Sales().setDate(LocalDate.of(2023, 1, 1))
                    .setSold(50)
                    .setPricePerUnit(10d),
            new Sales().setDate(LocalDate.of(2023, 1, 2))
                    .setSold(40)
                    .setPricePerUnit(11d),
            new Sales().setDate(LocalDate.of(2023, 1, 3))
                    .setSold(55)
                    .setPricePerUnit(9d)
    );

    try (Workbook workbook = excelMapper.createWorkbookFromObject(sales);
         var fileOutputStream = new FileOutputStream(fileName)) {
        workbook.write(fileOutputStream);
    }
    Assertions.assertThat(new File(fileName)).exists();
    Files.delete(Path.of(fileName));
}
```

# Excel to data classes example:

```java
    void readExcel() throws IOException {
        InputStream resourceAsStream = getClass().getClassLoader().getResourceAsStream("students.xlsx");
        assertThat(resourceAsStream).isNotNull();
        Workbook workbook = new XSSFWorkbook(resourceAsStream);

        ExcelMapper excelMapper = new ExcelMapperImpl();
        Map<String, List<Student>> objs = excelMapper.mapWorkbookToObjs(workbook, Student.class);
        assertThat(objs).isNotNull().size().isEqualTo(3);
        assertThat(objs.keySet()).contains("Class A").contains("Class B").contains("Class C");
        objs.forEach((s, students) -> {
                    assertThat(students).isNotNull().size().isEqualTo(4);
                }
        );
    }
```