package hckrn.simpleexcelmapper.exception;

public class ExcelMapperException extends RuntimeException {
    public ExcelMapperException(String message) {
        super(message);
    }

    public ExcelMapperException(Throwable throwable) {
        super(throwable);
    }
}
