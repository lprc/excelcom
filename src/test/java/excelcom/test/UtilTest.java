package excelcom.test;

import excelcom.util.Util;
import org.junit.Test;
import static org.junit.Assert.*;


/**
 * Tests for Util Class
 */
public class UtilTest {

    @Test
    public void shouldGetCharacterPositionInAlphabet() {
        assertEquals(1, Util.getPositionInAlphabet('a'));
        assertEquals(1, Util.getPositionInAlphabet('A'));
        assertEquals(4, Util.getPositionInAlphabet('d'));
        assertEquals(4, Util.getPositionInAlphabet('D'));
        assertEquals(26, Util.getPositionInAlphabet('z'));
        assertEquals(26, Util.getPositionInAlphabet('Z'));
    }

    @Test
    public void shouldGetRangeSize() {
        assertArrayEquals(new int[]{1,1}, Util.getRangeSize("C3"));
        assertArrayEquals(new int[]{2,3}, Util.getRangeSize("A1:C2"));
        assertArrayEquals(new int[]{3,3}, Util.getRangeSize("AA1:AC3"));
        assertArrayEquals(new int[]{3,3}, Util.getRangeSize("AAA1:AAC3"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void shouldNoticeWrongRangeFormat() {
        Util.getRangeSize("A2:B5:C6");
    }

    @Test(expected = IllegalArgumentException.class)
    public void shouldNoticeWrongColumnSize() {
        Util.getRangeSize("ABCD3:ABCDE4");
    }

    @Test(expected = IllegalArgumentException.class)
    public void shouldNoticeWrongRowSize() {
        Util.getRangeSize("A2000000:A2000001");
    }
}
