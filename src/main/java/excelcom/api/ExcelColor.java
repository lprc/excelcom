package excelcom.api;

/**
 * Excel colors
 */
public enum ExcelColor {
    NO_FILL(0),
    XL_NONE(-4142),
    BLACK(1),
    WHITE(2),
    RED(3),
    DARK_RED(9),
    PINK(7),
    ROSE(38),
    BROWN(30),
    ORANGE(46),
    LIGHT_ORANGE(45),
    GOLD(44),
    TAN(40),
    OLIVE_GREEN(52),
    DARK_YELLOW(12),
    LIME(43),
    YELLOW(6),
    LIGHT_YELLOW(36),
    DARK_GREEN(51),
    GREEN(10),
    SEA_GREEN(50),
    BRIGHT_GREEN(4),
    LIGHT_GREEN(35),
    DARK_TEAL(49),
    TEAL(14),
    AQUA(42),
    TURQUOISE(8),
    LIGHT_TURQUOISE(20),
    DARK_BLUE(11),
    BLUE(5),
    LIGHT_BLUE(41),
    SKY_BLUE(33),
    PALE_BLUE(37),
    INDIGO(55),
    BLUE_GRAY(47),
    VIOLET(13),
    PLUM(18),
    LAVENDER(39),
    GRAY_80(56),
    GRAY_50(16),
    GRAY_40(48),
    GRAY_25(15);

    private int index;

    private ExcelColor(int index) {
        this.index = index;
    }

    public int getIndex() {
        return index;
    }

    public static ExcelColor getColor(int index) {
        for(ExcelColor c : values()) {
            if(c.getIndex() == index) {
                return c;
            }
        }
        throw new IllegalArgumentException("no color with index " + index);
    }
}
