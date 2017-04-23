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

    private ExcelConnection connection = null;
    private File file = null;

    Workbook(IDispatch iDispatch) throws COMException {
        super(iDispatch);
    }

    Workbook(IDispatch iDispatch, ExcelConnection connection, File file) throws COMException {
        super(iDispatch);
        this.connection = connection;
        this.file = file;
    }

    /**
     * Sets a new name for the workbook
     * @param name Name to be set
     */
    public void setName(String name) {
        this.setProperty("Name", name);
    }

    /**
     * Gets the name of the workbook
     * @return Name of the workbook
     */
    public String getName() {
        return this.getStringProperty("Name");
    }

    /**
     * Closes the workbook
     * @param save true if changes should be saved
     */
    public void close(boolean save) {
        this.invokeNoReply("Close", new Variant.VARIANT(save));
    }

    /**
     * Saves this workbook
     * @throws COMException
     */
    public void save() throws COMException {
        this.invokeNoReply("Save");
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
    public Worksheet addWorksheet(String name) {
        return getWorksheets().addWorksheet(name);
    }

    /**
     * Gets the named worksheet
     * @param name Name of worksheet to get
     * @return excelcom.api.Worksheet
     */
    public Worksheet getWorksheet(String name) {
        return new Worksheet(this.getAutomationProperty("Worksheets", this, new Variant.VARIANT(name)));
    }
}
