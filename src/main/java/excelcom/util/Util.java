package excelcom.util;

import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.WTypes;

import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * A few simple utility functions
 */
public class Util {

    /**
     * Transposes the matrix
     * @param matrix Matrix to be transposed
     * @return a transposed copy of the matrix
     */
    public static Object[][] transpose(Object[][] matrix) {
        Object[][] temp = new Object[matrix[0].length][matrix.length];
        for (int i = 0; i < matrix.length; i++)
            for (int j = 0; j < matrix[0].length; j++)
                temp[j][i] = matrix[i][j];
        return temp;
    }

    /**
     * Checks if the string contains any special characters
     * @param string string to be tested
     * @return true if string contains any special character
     */
    public static boolean containsSpecialCharacters(String string) {
        Pattern p = Pattern.compile("[^a-z0-9 ]", Pattern.CASE_INSENSITIVE);
        return p.matcher(string).find();
    }

    /**
     * Prints a matrix to console
     * @param matrix Matrix to be printed
     * @return Printed string
     */
    public static String printMatrix(Object[][] matrix) {
        StringBuilder builder = new StringBuilder("");
        int rowCount = matrix.length;

        if(rowCount == 0) {
            System.out.println("Matrix is empty");
            return "Matrix is empty";
        }
        int columnCount = matrix[0].length;

        for(int i = 0; i < rowCount; i++) {
            builder.append("[");
            for(int j = 0; j < columnCount; j++) {
                builder.append(matrix[i][j]).append(j == columnCount - 1 ? "" : ",");
            }
            builder.append("]\n");
        }
        System.out.println(builder.toString());
        return builder.toString();
    }

    /**
     * Returns the size of the range. [rows, columns]
     * @param range range to be parsed, assuming it's like A3:B10 or AA10:CB20 or A13
     * @return integer array with two elements: [row, columns]
     */
    public static int[] getRangeSize(String range) throws IllegalArgumentException {
        String[] splits = range.split("[:]");
        if(splits.length == 1) {
            // unary range
            return new int[]{1,1};
        } else if(splits.length == 2) {
            // range is a matrix
            Matcher mFromLetters = Pattern.compile("[a-zA-Z]+").matcher(splits[0]);
            Matcher mFromDigits = Pattern.compile("[0-9]+").matcher(splits[0]);
            Matcher mToLetters = Pattern.compile("[a-zA-Z]+").matcher(splits[1]);
            Matcher mToDigits = Pattern.compile("[0-9]+").matcher(splits[1]);

            if(mFromLetters.find() && mFromDigits.find() && mToLetters.find() && mToDigits.find()) {
                String fromLetters = mFromLetters.group();
                int fromDigits = Integer.parseInt(mFromDigits.group());
                String toLetters = mToLetters.group();
                int toDigits = Integer.parseInt(mToDigits.group());

                if(fromLetters.length() > 3 || toLetters.length() > 3 || fromDigits > 1048576 || toDigits > 1048576) {
                    throw new IllegalArgumentException("range too big: " + range);
                } else if(fromDigits > toDigits) {
                    throw new IllegalArgumentException("begin of range is bigger than it's end");
                } else {
                    int beginColumn = fromLetters.length() != 3 ? (fromLetters.length() != 2 ? getPositionInAlphabet(fromLetters.charAt(0))
                            : getPositionInAlphabet(fromLetters.charAt(1)) + getPositionInAlphabet(fromLetters.charAt(0)) * 26)
                            : getPositionInAlphabet(fromLetters.charAt(2)) + getPositionInAlphabet(fromLetters.charAt(1)) * 26 + getPositionInAlphabet(fromLetters.charAt(0)) * 26 * 26;
                    int endColumn = toLetters.length() != 3 ? (toLetters.length() != 2 ? getPositionInAlphabet(toLetters.charAt(0))
                            : getPositionInAlphabet(toLetters.charAt(1)) + getPositionInAlphabet(toLetters.charAt(0)) * 26)
                            : getPositionInAlphabet(toLetters.charAt(2)) + getPositionInAlphabet(toLetters.charAt(1)) * 26 + getPositionInAlphabet(toLetters.charAt(0)) * 26 * 26;
                    return new int[]{(toDigits - fromDigits) + 1, (endColumn - beginColumn) + 1};
                }
            } else {
                throw new IllegalArgumentException("Unknown range format: " + range);
            }
        } else {
            throw new IllegalArgumentException("Unknown range format: " + range);
        }
    }

    /**
     * Gets the position of a character in alphabet
     * @param c character to be checked
     * @return position of character in alphabet
     */
    public static int getPositionInAlphabet(char c) {
        if(Character.isUpperCase(c)) {
            return ((int)c) - 64;
        } else {
            return ((int)c) - 96;
        }
    }

    /**
     * Gets the letter representation of a column, like in excel.
     * E.g. <code>getColumnName(1) => 'A' , getColumnName(27) => 'AA'</code>
     * @param column xolumn index, 1-based
     * @return column as letter
     */
    public static String getColumnName(int column) {
        if (column < 1 || column > 16384) {
            throw new IllegalArgumentException("column index is out of range (1 to 16384)");
        }
        int dividend = column;
        String columnName = "";
        int modulo;

        while (dividend > 0)
        {
            modulo = (dividend - 1) % 26;
            columnName = ((char)(65 + modulo)) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }

    /**
     * Converts bounds to range
     * @param lowerRow lower row bound, 0-based
     * @param lowerColumn lower row bound, 0-based
     * @param upperRow lower row bound, 0-based
     * @param upperColumn lower row bound, 0-based
     * @return range as string
     */
    public static String boundsToRange(int lowerRow, int lowerColumn, int upperRow, int upperColumn) {
        return Util.getColumnName(lowerColumn + 1) + Integer.toString(lowerRow + 1) + ":"
                + Util.getColumnName(upperColumn + 1) + Integer.toString(upperRow + 1);
    }

    /**
     * Creates a new Variant object which encapsulates the raw Object with correct type
     * @param o object to be wrapped
     * @return Variant wrapping object
     * @throws COMException if type is unknown
     */
    public static Variant.VARIANT createVariantFromObject(Object o) throws COMException {
        if (o instanceof String) {
            if (Util.containsSpecialCharacters((String) o)) {
                // insert strings with special characters as BSTR
                return new Variant.VARIANT(new WTypes.BSTR((String) o));
            } else {
                return new Variant.VARIANT((String) o);
            }
        } else if (o instanceof Integer) return new Variant.VARIANT((Integer) o);
        else if (o instanceof Float) return new Variant.VARIANT((Float) o);
        else if (o instanceof Double) return new Variant.VARIANT((Double) o);
        else if (o instanceof Long) return new Variant.VARIANT((Long) o);
        else if (o instanceof Date) return new Variant.VARIANT((Date) o);
        else if (o instanceof Short) return new Variant.VARIANT((Short) o);
        else if (o instanceof Boolean) return new Variant.VARIANT((Boolean) o);
        else if (o instanceof Byte) return new Variant.VARIANT((Byte) o);
        else if (o == null) return new Variant.VARIANT();
        throw new COMException("unknown type when creating variant from object: " + o.getClass());
    }
}
