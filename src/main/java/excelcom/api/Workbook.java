package excelcom.api;

import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.COMLateBindingObject;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant;

import java.io.File;

/**
 * Represents a excelcom.api.Workbook
 */
public class Workbook extends COMLateBindingObject {

    Workbook(IDispatch iDispatch) throws COMException {
        super(iDispatch);
    }

    /**
     * Gets the name of the workbook
     * @return Name of the workbook
     */
    public String getName() {
        try {
            return this.getStringProperty("Name");
        } catch (COMException e) {
            throw new ExcelException(e, "Failed get name of workbook");
        }
    }

    /**
     * Closes the workbook
     * @param save true if changes should be saved
     */
    public void close(boolean save) {
        try {
            this.invokeNoReply("Close", new Variant.VARIANT(save));
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to " + (save ? "save and " : "") + "close workbook");
        }
    }

    /**
     * Saves this workbook
     * @throws COMException
     */
    public void save() throws COMException {
        try {
            this.invokeNoReply("Save");
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to save workbook");
        }
    }

    /**
     * Saves this workbook to a new file
     * @param file new file
     * @throws ExcelException if saving fails
     */
    public void saveAs(File file) throws ExcelException {
        try {
            this.invokeNoReply("SaveAs", new Variant.VARIANT(file.getAbsolutePath()));
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to save workbook to " + file.getAbsolutePath());
        }
    }

    /**
     * Gets a list of worksheets in this workbook
     * @return list of worksheets
     */
    public Worksheets getWorksheets() {
        return new Worksheets(this.getAutomationProperty("Worksheets"));
    }

    /**
     * Adds a worksheet to this workbook
     * @return a excelcom.api.Worksheet instance representing the newly created worksheet
     */
    public Worksheet addWorksheet(String name) throws ExcelException {
        try {
            return getWorksheets().addWorksheet(name);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to add worksheet named " + name);
        }
    }

    /**
     * Gets the named worksheet
     * @param name Name of worksheet to get
     * @return excelcom.api.Worksheet
     */
    public Worksheet getWorksheet(String name) throws ExcelException {
        try {
            return new Worksheet(this.getAutomationProperty("Worksheets", this, new Variant.VARIANT(name)));
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get worksheet named " + name);
        }
    }
}
