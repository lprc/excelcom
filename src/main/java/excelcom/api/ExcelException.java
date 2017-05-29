package excelcom.api;

import com.sun.jna.platform.win32.COM.COMException;

/**
 * An exception thrown by COM using Excel
 * */
public class ExcelException extends COMException {
    ExcelException(COMException e, String message) {
        super(message + "\n" + e.getMessage());
    }
}
