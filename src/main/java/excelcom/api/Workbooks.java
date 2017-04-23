package excelcom.api;

import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.COMLateBindingObject;
import com.sun.jna.platform.win32.COM.IDispatch;

/**
 * Represents a List of Workbooks as COM object
 * Only for internal use
 */
class Workbooks extends COMLateBindingObject {
    Workbooks(IDispatch iDispatch) throws COMException {
        super(iDispatch);
    }
}
