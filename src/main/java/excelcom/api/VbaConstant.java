package excelcom.api;

/**
 * VBA Constants
 */
public enum VbaConstant {
    XL_FORMULAS(-4123),
    XL_VALUES(-4163),
    XL_NOTES(-4144),
    XL_WHOLE(1),
    XL_PART(2),
    XL_BY_ROWS(1),
    XL_BY_COLUMNS(2),
    XL_NEXT(1),
    XL_PREVIOUS(2);

    private int index;

    private VbaConstant(int index) {
        this.index = index;
    }

    public int getIndex() {
        return index;
    }
}
