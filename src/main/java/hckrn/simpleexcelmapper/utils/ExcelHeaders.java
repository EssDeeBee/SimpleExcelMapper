package hckrn.simpleexcelmapper.utils;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.HashMap;

@Data
@AllArgsConstructor
public class ExcelHeaders {
    private HashMap<String, Integer> headersIndexes;
    private int headersRowNumber;
}
