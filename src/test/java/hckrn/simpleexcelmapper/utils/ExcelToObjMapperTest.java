package hckrn.simpleexcelmapper.utils;

import hckrn.simpleexcelmapper.utils.dto.Student;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class ExcelToObjMapperTest {

    @Test
    void shouldMapExcelToObjectsWhenExcelIsProvided() throws IOException {
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
}