package hckrn.simpleexcelmapper.utils;

import hckrn.simpleexcelmapper.utils.dto.Sales;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Workbook;
import org.assertj.core.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

class SalesMappingTest {

    @Test
    @SneakyThrows
    void shouldCreateSalesReport() {
        var excelMapper = new ExcelMapperImpl();

        var fileName = "sales-out-" + LocalTime.now().format(DateTimeFormatter.ofPattern("hh.mm.ss")) + ".xlsx";
        List<Sales> sales = getSales();

        try (Workbook workbook = excelMapper.createWorkbookFromObject(sales);
             var fileOutputStream = new FileOutputStream(fileName)) {
            workbook.write(fileOutputStream);
        }
        Assertions.assertThat(new File(fileName)).exists();
        Files.delete(Path.of(fileName));
    }

    private List<Sales> getSales() {
        return List.of(
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

    }
}
