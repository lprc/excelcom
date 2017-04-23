package excelcom.api;

import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.COMLateBindingObject;
import com.sun.jna.platform.win32.COM.IDispatch;

/**
 * List of worksheets as COM object
 * Only for internal use
 */
class Worksheets extends COMLateBindingObject {
    Worksheets(IDispatch iDispatch) throws COMException {
        super(iDispatch);
    }

    /**
     * Adds a new excelcom.api.Worksheet with the given name
     * @param name Name of new worksheet
     */
    Worksheet addWorksheet(String name) {
        Worksheet ws = new Worksheet(this.getAutomationProperty("Add"));
        ws.setName(name);
        return ws;
    }
}
