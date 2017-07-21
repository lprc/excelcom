package excelcom.api;

import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.WTypes;

/**
 * Represents a find result object. Row and Column number of the cell can be queried from it.
 */
public class FindResult extends Range {
    Range searchedRange;
    int row = -1, column = -1;

    FindResult(IDispatch iDispatch, Range searchedRange) {
        super(iDispatch);
        this.searchedRange = searchedRange;
        this.row = getRow();
        this.column = getColumn();
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        if (!super.equals(o)) return false;

        FindResult that = (FindResult) o;

        if (row != that.row) return false;
        return column == that.column;
    }

    @Override
    public int hashCode() {
        int result = super.hashCode();
        result = 31 * result + row;
        result = 31 * result + column;
        return result;
    }

    @Override
    public String toString() {
        return "(" + getRow() + "," + getColumn() + ") -> " + getContent();
    }

    public int getRow() {
        if(row == -1) {
           row = this.invoke("Row").intValue() - 1;
        }
        return row;
    }

    public int getColumn() {
        if(column == -1) {
            column = this.invoke("Column").intValue() - 1;
        }
        return column;
    }

    /**
     * Gets the content of the cell
     * @return content of cell which was found
     */
    public Object getContent() throws ExcelException {
        try {
            Object val = this.invoke("Value").getValue();
            return (val instanceof WTypes.BSTR) ? val.toString() : val;
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to content from find result");
        }
    }

    /**
     * Gets the next occurrence of the value searched before (or the same result if there is only one occurrence)
     * @return result of search
     */
    public FindResult next() throws ExcelException {
        try {
            FindResult fr = searchedRange.findNext(this);
            if (fr.getColumn() == this.getColumn() && fr.getRow() == this.getRow()) {
                return this;
            } else {
                return fr;
            }
        } catch (COMException e) {
            throw new ExcelException(e, "Failed to get next find result");
        }
    }


}
