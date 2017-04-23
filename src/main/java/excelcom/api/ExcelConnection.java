package excelcom.api;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.COMLateBindingObject;
import com.sun.jna.platform.win32.Ole32;
import com.sun.jna.platform.win32.Variant;

import java.io.File;

/**
 * Represents a connection to an excel instance
 */
public class ExcelConnection extends COMLateBindingObject {

    /**
     * Connects to a new excel instance
     * @return excel connection
     * @throws COMException when connecting fails
     */
    public static ExcelConnection connect() throws COMException {
        return connect(false);
    }

    /**
     * Connects to a new excel instance
     * @param useActiveInstance if true, an existing instance will be used
     * @return excel connection
     * @throws COMException when connecting fails
     */
    public static ExcelConnection connect(boolean useActiveInstance) throws COMException {
        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
        Runtime.getRuntime().addShutdownHook(new Thread() {
            @Override
            public void run() {
                Ole32.INSTANCE.CoUninitialize();
            }
        });
        return new ExcelConnection(useActiveInstance);
    }

    /**
     * Connects to an excel instance
     * @param useActiveInstance true if should connect to an active excel instance
     */
    private ExcelConnection(boolean useActiveInstance) throws COMException {
        super("Excel.Application", useActiveInstance);
    }

    /* ****************************
     * Connection specific methods
     * ****************************/
    /**
     * Show or hide the excel instance
     * @param bVisible true if excel instance should be shown
     * @throws COMException
     */
    public void setVisible(boolean bVisible) throws COMException {
        this.setProperty("Visible", bVisible);
    }

    public void setDisplayAlerts(boolean displayAlerts) {
        this.setProperty("DisplayAlerts", displayAlerts);
    }

    public String getVersion() throws COMException {
        return this.getStringProperty("Version");
    }

    public void quit() throws COMException {
        this.invokeNoReply("Quit");
    }

    /* **************************
     * Workbook specific methods
     * **************************/

    /**
     * @return list of workbooks opened in this excel instance
     */
    public Workbooks getWorkbooks() {
        return new Workbooks(this.getAutomationProperty("WorkBooks"));
    }

    /**
     * Gets the currently active workbook
     * @return Currently active excelcom.api.Workbook
     */
    public Workbook getActiveWorkbook() {
        return new Workbook(this.getAutomationProperty("ActiveWorkbook"));
    }

    /**
     * Opens a workbook
     * @param file file to open
     * @return excelcom.api.Workbook instance
     */
    public Workbook openWorkbook(File file) {
        this.invokeNoReply("Open", getWorkbooks(), new Variant.VARIANT(file.getAbsolutePath()));
        return this.getActiveWorkbook(); //TODO application.recentfiles(1)
    }
}
