package hckrn.simpleexcelmapper.utils;

import hckrn.simpleexcelmapper.utils.dto.Student;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.SecureRandom;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class ExcelUtilsServiceTest {
    private final SecureRandom secureRandom = new SecureRandom();

    @Test
    void shouldMapObjToExcel() throws IOException {

        String fileName = "students-out-" + LocalTime.now().format(DateTimeFormatter.ofPattern("hh.mm.ss")) + ".xlsx";
        List<Student> students = createStudents();

        Workbook workbookFromObject = new ExcelUtilsService().createWorkbookFromObject(students);
        try (var fileOutputStream = new FileOutputStream(fileName)) {
            workbookFromObject.write(fileOutputStream);
        }

        try (var fileInputStream = new FileInputStream(fileName)) {
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            var excelMapper = new ExcelMapper();
            Map<String, List<Student>> objs = excelMapper.mapWorkbookToObjs(workbook, Student.class);
            assertThat(objs).isNotNull().size().isEqualTo(1);
            objs.forEach((s, sts) -> assertThat(sts).isNotNull().size().isEqualTo(students.size() + 1));
        }

        Files.delete(Path.of(fileName));
    }

    private List<Student> createStudents() {
        var students = new LinkedList<Student>();

        for (int i = 0; i < 10; i++) {
            Student student = new Student();
            student.setStudentId(i + 1);
            student.setName(getRandomName());
            student.setAge(secureRandom.nextInt(16, 40));
            student.setAdmissionDate(LocalDate.now().minusMonths(secureRandom.nextInt(0, 5)));
            students.add(student);
        }
        return students;
    }

    private String getRandomName() {
        var names = List.of(
                "Michael Johnson",
                "Emily Roberts",
                "Jordan Taylor",
                "Jessica Martinez",
                "Brian Anderson",
                "Ashley Thompson",
                "Brandon White",
                "Nicole Clark",
                "Tyler Wilson",
                "Rachel Garcia",
                "Dylan Hernandez",
                "Rebecca Lee",
                "Alex Smith",
                "Chelsea Jones",
                "Christopher Brown"
        );

        return names.get(secureRandom.nextInt(0, names.size()));

    }

}