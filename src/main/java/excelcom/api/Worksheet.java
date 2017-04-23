package excelcom.api;

import com.sun.jna.platform.win32.*;
import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.COMLateBindingObject;
import com.sun.jna.platform.win32.COM.IDispatch;
import excelcom.util.Util;

import java.util.Date;
import java.util.regex.Pattern;

/**
 * Represents a worksheet
 */
public class Worksheet extends COMLateBindingObject {
    public Worksheet(IDispatch iDispatch) {
        super(iDispatch);
    }

    /**
     * Sets a new name for the worksheet
     * @param name Name to be set
     */
    public void setName(String name) {
        this.setProperty("Name", name);
    }

    /**
     * Gets the name of the worksheet
     * @return Name of the worksheet
     */
    public String getName() {
        return this.getStringProperty("Name");
    }

    /**
     * Gets the whole used content of the worksheet (using UsedRange)
     * @return 2-dimensional Array of Object with content
     * @throws COMException
     */
    public Object[][] getContent() throws COMException {
        return this.getContent("UsedRange");
    }

    /**
     * Gets the content in range
     * @param range Range with content
     * @return 2-dimensional array of content. If an unary range was given, the element is in result[0][0].
     * @throws COMException
     * //todo class for return values (?)
     */
    public Object[][] getContent(String range) throws COMException {
        Range pRange = range.equals("UsedRange") ?
                new Range(this.getAutomationProperty("UsedRange", this)) :
                new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        Object contentRaw = pRange.getValue().getValue();

        if(contentRaw instanceof OaIdl.SAFEARRAY) {
            // TODO Workaround for bug in toPrimitiveArray (0-based java vs 1-based excel) see github issue #785 (https://github.com/java-native-access/jna/issues/785)
            ((OaIdl.SAFEARRAY) contentRaw).rgsabound[0].lLbound = new WinDef.LONG(0);
            ((OaIdl.SAFEARRAY) contentRaw).rgsabound[1].lLbound = new WinDef.LONG(0);
            return Util.transpose((Object[][]) OaIdlUtil.toPrimitiveArray((OaIdl.SAFEARRAY)contentRaw, true));
        } else {
            return new Object[][]{{contentRaw}};
        }
    }

    /**
     * Sets the content of range
     * @param range range with content to be set
     * @param content content to be set
     */
    public void setContent(String range, Object[][] content) {
        Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        int rowCount = content.length;
        int columnCount = content[0].length;

        // transpose content: in java it's (row,column) but in excel it's (column,row)
        OaIdl.SAFEARRAY sa = OaIdl.SAFEARRAY.createSafeArray(columnCount, rowCount);
        for(int i = 0; i < rowCount; i++) {
            for(int j = 0; j < columnCount; j++) {
                Object cell = content[i][j];
                if(cell instanceof String) {
                    if(Util.containsSpecialCharacters((String)cell)) {
                        // insert strings with special characters as BSTR
                        sa.putElement(new Variant.VARIANT(new WTypes.BSTR((String)cell)), j ,i);
                    } else {
                        sa.putElement(new Variant.VARIANT((String)cell), j ,i);
                    }
                }
                else if(cell instanceof Integer) sa.putElement(new Variant.VARIANT((Integer)cell), j ,i);
                else if(cell instanceof Float) sa.putElement(new Variant.VARIANT((Float)cell), j ,i);
                else if(cell instanceof Double) sa.putElement(new Variant.VARIANT((Double)cell), j ,i);
                else if(cell instanceof Long) sa.putElement(new Variant.VARIANT((Long)cell), j ,i);
                else if(cell instanceof Date) sa.putElement(new Variant.VARIANT((Date)cell), j ,i);
                else if(cell instanceof Short) sa.putElement(new Variant.VARIANT((Short)cell), j ,i);
                else if(cell instanceof Boolean) sa.putElement(new Variant.VARIANT((Boolean)cell), j ,i);
                else if(cell instanceof Byte) sa.putElement(new Variant.VARIANT((Byte)cell), j ,i);
            }
        }

        // set content
        this.setProperty("Value", pRange, new Variant.VARIANT(sa));
    }

    /**
     * Sets the content of a range to one value
     * @param range range
     * @param content value to be set
     */
    public void setContent(String range, Object content) {
        int[] rangeSize = Util.getRangeSize(range);
        Object[][] temp = new Object[rangeSize[0]][rangeSize[1]];
        // set content to each cell in range
        for(int row = 0; row < rangeSize[0]; row++) {
            for(int column = 0; column < rangeSize[1]; column++) {
                temp[row][column] = content;
            }
        }
        setContent(range, temp);
    }

    /**
     * Sets the fill (background) color of a range
     */
    public void setFillColor(String range, ExcelColor color) {
        Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        pRange.setInteriorColor(color);
    }

    /**
     * Gets the fill color of the range
     * @throws NullPointerException if range has multiple fill colors (or an unexpected error appears)
     */
    public ExcelColor getFillColor(String range) throws NullPointerException {
        Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        return pRange.getInteriorColor();
    }

    /**
     * Sets the font color of a range
     */
    public void setFontColor(String range, ExcelColor color) {
        Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        pRange.setFontColor(color);
    }

    /**
     * Gets the font color of the range
     * @throws NullPointerException if range has multiple fill colors (or an unexpected error appears)
     */
    public ExcelColor getFontColor(String range) throws NullPointerException {
        Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        return pRange.getFontColor();
    }

    /**
     * Sets the border color of a range
     */
    public void setBorderColor(String range, ExcelColor color) {
        Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        pRange.setBorderColor(color);
    }

    /**
     * Gets the border color of the range
     * @throws NullPointerException if range has multiple fill colors (or an unexpected error appears)
     */
    public ExcelColor getBorderColor(String range) throws NullPointerException {
        Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        return pRange.getBorderColor();
    }

    /**
     * Sets a comment for one cell. Setting columns for multiple cells is not supported.
     * @param range cell
     * @param comment text
     * @throws IllegalArgumentException when a multiple cell range is given
     */
    public void setComment(String range, String comment) throws IllegalArgumentException {
        int[] rangeBounds = Util.getRangeSize(range);
        if(rangeBounds[0] != 1 || rangeBounds[1] != 1) {
            throw new IllegalArgumentException("multiple cell range given. comment can be set for one cell only.");
        }
        Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        pRange.setComment(comment);
    }

    /**
     * Gets the comment of a cell. Multiple cell ranges are not supported.
     * @param range cell
     * @return comment text
     * @throws IllegalArgumentException when a multiple cell range is given
     */
    public String getComment(String range) throws IllegalArgumentException {
        int[] rangeBounds = Util.getRangeSize(range);
        if(rangeBounds[0] != 1 || rangeBounds[1] != 1) {
            throw new IllegalArgumentException("multiple cell range given. comment can be read from one cell only.");
        }
        Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
        return pRange.getComment();
    }
}
