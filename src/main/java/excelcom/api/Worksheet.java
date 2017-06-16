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
    public void setName(String name) throws ExcelException {
        try {
            this.setProperty("Name", name);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set name of worksheet to " + name);
        }
    }

    /**
     * Gets the name of the worksheet
     * @return Name of the worksheet
     */
    public String getName() throws ExcelException {
        try {
            return this.getStringProperty("Name");
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get name of worksheet");
        }
    }

    /**
     * Deletes this worksheet. Displays a confirmation mesage box unless ExcelConnection#displayAlerts was set to false.
     * @return true if sheet was deleted, false otherwise
     * @throws ExcelException
     */
    public boolean delete() throws ExcelException {
        try {
            return this.invoke("Delete").booleanValue();
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to delete worksheet");
        }
    }

    /**
     * Gets content from one cell as an object
     * @param range one cell range, e.g. "A5"
     * @return cell value
     * @throws ExcelException
     * @throws IllegalArgumentException if multiple cell range was given
     */
    public Object getUnaryContent(String range) throws ExcelException, IllegalArgumentException {
        int[] rangeBounds = Util.getRangeSize(range);
        if (rangeBounds[0] != 1 || rangeBounds[1] != 1) {
            throw new IllegalArgumentException("Failed to get content from one cell. Multiple cell range was given: " + range);
        }
        try {
            Object c = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)))
                    .getValue().getValue();
            if(c instanceof WTypes.BSTR) {
                return c.toString();
            } else {
                return c;
            }
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get unary content in range '" + range + "'");
        }
    }

    /**
     * Sets content of one cell
     * @param range one cell range, e.g. "A5"
     * @param content
     * @throws IllegalArgumentException if multiple cell range was given
     * @throws ExcelException
     */
    public void setUnaryContent(String range, Object content) throws ExcelException, IllegalArgumentException {
        int[] rangeBounds = Util.getRangeSize(range);
        if (rangeBounds[0] != 1 || rangeBounds[1] != 1) {
            throw new IllegalArgumentException("Failed to set content of one cell to " + content + ". Multiple cell range was given: " + range);
        }
        try {
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            this.setProperty("Value", pRange, Util.createVariantFromObject(content));
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set unary content in range '" + range + "'");
        }
    }

    /**
     * Gets the whole used content of the worksheet (using UsedRange)
     * @return 2-dimensional Array of Object with content
     * @throws ExcelException
     */
    public Object[][] getContent() throws ExcelException {
        return this.getContent("UsedRange");
    }

    /**
     * Gets the content in range
     * @param range Range with content
     * @return 2-dimensional array of content. If an unary range was given, the element is in result[0][0].
     * @throws ExcelException
     */
    public Object[][] getContent(String range) throws ExcelException {
        try {
            Range pRange = range.equals("UsedRange") ?
                    new Range(this.getAutomationProperty("UsedRange", this)) :
                    new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            Object contentRaw = pRange.getValue().getValue();

            if (contentRaw instanceof OaIdl.SAFEARRAY) {
                // TODO Workaround for bug in toPrimitiveArray (0-based java vs 1-based excel) see github issue #785 (https://github.com/java-native-access/jna/issues/785)
                ((OaIdl.SAFEARRAY) contentRaw).rgsabound[0].lLbound = new WinDef.LONG(0);
                ((OaIdl.SAFEARRAY) contentRaw).rgsabound[1].lLbound = new WinDef.LONG(0);
                return Util.transpose((Object[][]) OaIdlUtil.toPrimitiveArray((OaIdl.SAFEARRAY) contentRaw, true));
            } else {
                return new Object[][]{{contentRaw}};
            }
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get content in range '" + range + "'");
        }
    }

    /**
     * Sets the content of range
     * @param range range with content to be set
     * @param content content to be set
     */
    public void setContent(String range, Object[][] content) throws ExcelException {
        try {
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            int rowCount = content.length;
            int columnCount = content[0].length;

            // transpose content: in java it's (row,column) but in excel it's (column,row)
            OaIdl.SAFEARRAY sa = OaIdl.SAFEARRAY.createSafeArray(columnCount, rowCount);
            for (int i = 0; i < rowCount; i++) {
                for (int j = 0; j < columnCount; j++) {
                    Object cell = content[i][j];
                    sa.putElement(Util.createVariantFromObject(cell), j, i);
                }
            }

            // set content
            this.setProperty("Value", pRange, new Variant.VARIANT(sa));
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set content in range '" + range + "'");
        }
    }

    /**
     * Sets the content of a range to one value
     * @param range range
     * @param content value to be set
     */
    public void setContent(String range, Object content) throws ExcelException {
        try {
            int[] rangeSize = Util.getRangeSize(range);
            Object[][] temp = new Object[rangeSize[0]][rangeSize[1]];
            // set content to each cell in range
            for (int row = 0; row < rangeSize[0]; row++) {
                for (int column = 0; column < rangeSize[1]; column++) {
                    temp[row][column] = content;
                }
            }
            setContent(range, temp);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set content in range '" + range + "'");
        }
    }

    /**
     * Sets the fill (background) color of a range
     */
    public void setFillColor(String range, ExcelColor color) throws ExcelException {
        try {
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            pRange.setInteriorColor(color);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set fill color in range '" + range + "'");
        }
    }

    /**
     * Gets the fill color of the range
     * @throws NullPointerException if range has multiple fill colors (or an unexpected error appears)
     */
    public ExcelColor getFillColor(String range) throws ExcelException, NullPointerException {
        try {
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            return pRange.getInteriorColor();
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get fill color in range '" + range + "'");
        }
    }

    /**
     * Sets the font color of a range
     */
    public void setFontColor(String range, ExcelColor color) throws ExcelException {
        try {
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            pRange.setFontColor(color);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set font color in range '" + range + "'");
        }
    }

    /**
     * Gets the font color of the range
     * @throws NullPointerException if range has multiple fill colors (or an unexpected error appears)
     */
    public ExcelColor getFontColor(String range) throws ExcelException, NullPointerException {
        try {
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            return pRange.getFontColor();
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get font color in range '" + range + "'");
        }
    }

    /**
     * Sets the border color of a range
     */
    public void setBorderColor(String range, ExcelColor color) throws ExcelException {
        try {
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            pRange.setBorderColor(color);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set border color in range '" + range + "'");
        }
    }

    /**
     * Gets the border color of the range
     * @throws NullPointerException if range has multiple fill colors (or an unexpected error appears)
     */
    public ExcelColor getBorderColor(String range) throws ExcelException, NullPointerException {
        try {
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            return pRange.getBorderColor();
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get border color in range '" + range + "'");
        }
    }

    /**
     * Sets a comment for one cell. Setting columns for multiple cells is not supported.
     * @param range cell
     * @param comment text
     * @throws IllegalArgumentException when a multiple cell range is given
     */
    public void setComment(String range, String comment) throws ExcelException, IllegalArgumentException {
        try {
            int[] rangeBounds = Util.getRangeSize(range);
            if (rangeBounds[0] != 1 || rangeBounds[1] != 1) {
                throw new IllegalArgumentException("multiple cell range given. comment can be set for one cell only.");
            }
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            pRange.setComment(comment);
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to set comment in range '" + range + "'");
        }
    }

    /**
     * Gets the comment of a cell. Multiple cell ranges are not supported.
     * @param range cell
     * @return comment text
     * @throws IllegalArgumentException when a multiple cell range is given
     */
    public String getComment(String range) throws ExcelException, IllegalArgumentException {
        try {
            int[] rangeBounds = Util.getRangeSize(range);
            if (rangeBounds[0] != 1 || rangeBounds[1] != 1) {
                throw new IllegalArgumentException("multiple cell range given. comment can be read from one cell only.");
            }
            Range pRange = new Range(this.getAutomationProperty("Range", this, new Variant.VARIANT(range)));
            return pRange.getComment();
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get comment in range '" + range + "'");
        }
    }
}
