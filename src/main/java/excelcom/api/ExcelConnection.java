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
     * @throws ExcelException when connecting fails
     */
    public static ExcelConnection connect() throws ExcelException {
        return connect(false);
    }

    /**
     * Connects to a new excel instance
     * @param useActiveInstance if true, an existing instance will be used
     * @return excel connection
     * @throws ExcelException when connecting fails
     */
    public static ExcelConnection connect(boolean useActiveInstance) throws ExcelException {
        try {
            Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
            Runtime.getRuntime().addShutdownHook(new Thread() {
                @Override
                public void run() {
                    Ole32.INSTANCE.CoUninitialize();
                }
            });
            return new ExcelConnection(useActiveInstance);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to connect to " + (useActiveInstance ? "an active " : "a new ") + "Excel instance");
        }
    }

    /**
     * Initializes COM manually, NOT RECOMMMENDED! ExcelConnection::connect should initialize and uninitialize COM automatically.
     * However if this method is called, uninitializeCom must be called anywhen later!
     * @throws COMException if initialization fails
     */
    public static void initializeCom() throws COMException {
        Ole32.INSTANCE.CoInitializeEx(Pointer.NULL, Ole32.COINIT_MULTITHREADED);
    }

    /**
     * Uninitialize COM manually, NOT RECOMMENDED! Only use this if you used initializeCom before.
     * @throws COMException if uninitialization fails
     */
    public static void uninitializeCom() throws COMException {
        Ole32.INSTANCE.CoUninitialize();
    }

    /**
     * Connects to an excel instance
     * @param useActiveInstance true if should connect to an active excel instance
     */
    private ExcelConnection(boolean useActiveInstance) throws ExcelException {
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
    public void setVisible(boolean bVisible) throws ExcelException {
        try {
            this.setProperty("Visible", bVisible);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set Property 'Visible' to " + bVisible);
        }
    }

    public void setDisplayAlerts(boolean displayAlerts) throws ExcelException {
        try {
            this.setProperty("DisplayAlerts", displayAlerts);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set Property 'DisplayAlerts' to " + displayAlerts);
        }
    }

    public String getVersion() throws ExcelException {
        try {
            return this.getStringProperty("Version");
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get Property 'Version'");
        }
    }

    public void quit() throws ExcelException {
        try {
            this.invokeNoReply("Quit");
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to invoke 'Quit'");
        }
    }

    /* **************************
     * Workbook specific methods
     * **************************/

    /**
     * @return list of workbooks opened in this excel instance
     */
    public Workbooks getWorkbooks() throws ExcelException {
        try {
            return new Workbooks(this.getAutomationProperty("WorkBooks"));
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get Property 'Workbooks'");
        }
    }

    /**
     * Gets the currently active workbook
     * @return Currently active excelcom.api.Workbook
     */
    public Workbook getActiveWorkbook() throws ExcelException {
        try {
            return new Workbook(this.getAutomationProperty("ActiveWorkbook"));
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get Property 'ActiveWorkbook'");
        }
    }

    /**
     * Opens a workbook
     * @param file file to open
     * @return excelcom.api.Workbook instance
     */
    public Workbook openWorkbook(File file) throws ExcelException {
        try {
            this.invokeNoReply("Open", getWorkbooks(), new Variant.VARIANT(file.getAbsolutePath()));
            return this.getActiveWorkbook(); //TODO application.recentfiles(1)
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to open Workbook located at " + file.getAbsolutePath());
        }
    }
}
