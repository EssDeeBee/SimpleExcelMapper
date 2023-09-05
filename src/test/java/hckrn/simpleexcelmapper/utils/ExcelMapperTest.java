package hckrn.simpleexcelmapper.utils;

import hckrn.simpleexcelmapper.annotation.ColumnExcel;
import lombok.Data;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class ExcelMapperTest {

    @Test
    void shouldMapExcelToObjectsWhenExcelIsProvided() throws IOException {
        InputStream resourceAsStream = getClass().getClassLoader().getResourceAsStream("students.xlsx");
        assertThat(resourceAsStream).isNotNull();
        Workbook workbook = new XSSFWorkbook(resourceAsStream);

        var excelMapper = new ExcelMapper();
        Map<String, List<Student>> objs = excelMapper.mapWorkbookToObjs(workbook, Student.class);
        assertThat(objs).isNotNull().size().isEqualTo(3);
        assertThat(objs.keySet()).contains("Class A").contains("Class B").contains("Class C");
        objs.forEach((s, students) -> {
                    assertThat(students).isNotNull().size().isEqualTo(4);
                }
        );
    }

    @Data
    static class Student {
        @ColumnExcel(name = "Student ID", position = 0)
        private Integer studentId;

        @ColumnExcel(name = "Name", position = 1)
        private String name;

        @ColumnExcel(name = "Age", position = 2)
        private Integer age;

        public String toString() {
            return "Student(studentId=" + this.getStudentId() + ", name=" + this.getName() + ", age=" + this.getAge() + ")";
        }
    }
}