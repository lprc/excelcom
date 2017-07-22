package excelcom.api;

/**
 * Class for setting options for find method
 */
public class FindOptions {
    private String value = "*";
    private String range = "UsedRange";
    private String after = null;
    private VbaConstant lookIn = VbaConstant.XL_FORMULAS;
    private VbaConstant lookAt = VbaConstant.XL_PART;
    private VbaConstant searchOrder = VbaConstant.XL_BY_ROWS;
    private VbaConstant searchDirection = VbaConstant.XL_NEXT;
    private Boolean matchCase = false;
    private Boolean matchByte = false;

    public String getValue() {
        return value;
    }

    public FindOptions setValue(String value) {
        this.value = value;
        return this;
    }

    public String getRange() {
        return range;
    }

    public FindOptions setRange(String range) {
        this.range = range;
        return this;
    }

    public String getAfter() {
        return after;
    }

    public FindOptions setAfter(String after) {
        this.after = after;
        return this;
    }

    public VbaConstant getLookIn() {
        return lookIn;
    }

    public FindOptions setLookIn(VbaConstant lookIn) {
        this.lookIn = lookIn;
        return this;
    }

    public VbaConstant getLookAt() {
        return lookAt;
    }

    public FindOptions setLookAt(VbaConstant lookAt) {
        this.lookAt = lookAt;
        return this;
    }

    public VbaConstant getSearchOrder() {
        return searchOrder;
    }

    public FindOptions setSearchOrder(VbaConstant searchOrder) {
        this.searchOrder = searchOrder;
        return this;
    }

    public VbaConstant getSearchDirection() {
        return searchDirection;
    }

    public FindOptions setSearchDirection(VbaConstant searchDirection) {
        this.searchDirection = searchDirection;
        return this;
    }

    public Boolean getMatchCase() {
        return matchCase;
    }

    public FindOptions setMatchCase(Boolean matchCase) {
        this.matchCase = matchCase;
        return this;
    }

    public Boolean getMatchByte() {
        return matchByte;
    }

    public FindOptions setMatchByte(Boolean matchByte) {
        this.matchByte = matchByte;
        return this;
    }
}
