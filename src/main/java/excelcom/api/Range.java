package excelcom.api;

import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.COMLateBindingObject;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant;

import static com.sun.jna.platform.win32.Variant.VT_NULL;

/**
 * Represents a Range
 */
class Range extends COMLateBindingObject {
    Range(IDispatch iDispatch) throws COMException {
        super(iDispatch);
    }

    Variant.VARIANT getValue() {
        return this.invoke("Value");
    }

    void setInteriorColor(ExcelColor color) {
        new CellPane(this.getAutomationProperty("Interior", this)).setColorIndex(color);
    }

    ExcelColor getInteriorColor() {
        return ExcelColor.getColor(new CellPane(this.getAutomationProperty("Interior", this)).getColorIndex());
    }

    void setFontColor(ExcelColor color) {
        new CellPane(this.getAutomationProperty("Font", this)).setColorIndex(color);
    }

    ExcelColor getFontColor() {
        return ExcelColor.getColor(new CellPane(this.getAutomationProperty("Font", this)).getColorIndex());
    }

    void setBorderColor(ExcelColor color) {
        new CellPane(this.getAutomationProperty("Borders", this)).setColorIndex(color);
    }

    ExcelColor getBorderColor() {
        return ExcelColor.getColor(new CellPane(this.getAutomationProperty("Borders", this)).getColorIndex());
    }

    void setComment(String comment) {
        this.invokeNoReply("ClearComments");
        this.invoke("AddComment", new Variant.VARIANT(comment));
    }

    String getComment() {
        return new COMLateBindingObject(this.getAutomationProperty("Comment")) {
            private String getText() {
                return this.invoke("Text").stringValue();
            }
        }.getText();
    }

    /**
     * Can be Interior, Border or Font. Has methods for setting e.g. Color.
     */
    private class CellPane extends COMLateBindingObject {
        CellPane(IDispatch iDispatch) {
            super(iDispatch);
        }

        void setColorIndex(ExcelColor color) {
            this.setProperty("ColorIndex", color.getIndex());
        }

        int getColorIndex() {
            Variant.VARIANT colorIndex = this.invoke("ColorIndex");
            if(colorIndex.getVarType().intValue() == VT_NULL) {
                throw new NullPointerException("return type of colorindex is null. Maybe multiple colors in range?");
            }
            return this.invoke("ColorIndex").intValue();
        }
    }
}
