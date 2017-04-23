package excelcom.test;

import excelcom.api.ExcelColor;
import excelcom.api.ExcelConnection;
import excelcom.api.Workbook;

import excelcom.api.Worksheet;
import excelcom.util.Util;
import org.junit.*;

import java.io.File;
import java.net.URISyntaxException;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import static org.junit.Assert.*;

/**
 * Test cases for excel
 */
public class ExcelTest {

    private ExcelConnection connection = null;
    private Workbook workbook = null;

    @Before
    public void establishConnection() throws URISyntaxException, InterruptedException {
        connection = ExcelConnection.connect();
        connection.setDisplayAlerts(false);
        workbook = connection.openWorkbook(new File(getClass().getResource("../../test.xlsx").toURI()));
    }

    @After
    public void closeConnections() {
        if(workbook != null) {
            workbook.close(false);
        }
        if(connection != null) {
            connection.quit();
        }
    }

    @Test
    public void shouldConnect() {
        assertNotNull(connection);
        assertNotNull(connection.getVersion());
    }

    @Test
    public void shouldOpenWorkbook() {
        assertNotNull(workbook);
    }

    @Test
    public void shouldAddWorksheet() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);
    }

    @Test
    public void shouldGetWorksheetByName() {
        Worksheet ws = null;
        Worksheet ws2 = null;

        ws = workbook.addWorksheet("test");
        assertNotNull(ws);
        ws2 = ws;
        ws = workbook.getWorksheet("test");
        assertEquals(ws2, ws);
    }


    @Test
    public void shouldModifyWorksheetContent() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);

        String range = "A2:B5";
        try {
            Object[][] content = new Object[][]{
                    {"A22", 123},
                    {54.6, 23.5f},
                    {"äöüß", "#?-"},
                    {new SimpleDateFormat("dd.MM.yyyy").parse("03.03.2017"), "=Summe(A3;B3)"}
            };

            ws.setContent(range, content);
            Util.printMatrix(ws.getContent(range));

            Object[][] actual = ws.getContent(range);
            assertEquals(content[0][0], actual[0][0]);
            assertEquals(content[0][1], ((Double)actual[0][1]).intValue());
            assertEquals(content[1][0], actual[1][0]);
            assertEquals(content[1][1], ((Double)actual[1][1]).floatValue());
            assertEquals(content[2][0], actual[2][0]);
            assertEquals(content[2][1], actual[2][1]);
            assertEquals(content[3][0], actual[3][0]);
            assertEquals(78.1, (Double)actual[3][1], 0.1);
        } catch (ParseException e) {
            fail(e.getClass() + ": " + e.getMessage());
        }
    }

    @Test
    public void shouldSetRangeToOneValue() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);

        String range = "A2:B5";
        Object content = 123.5;
        Object[][] expectedContent = new Object[][]{
                {123.5, 123.5},
                {123.5, 123.5},
                {123.5, 123.5},
                {123.5, 123.5}
        };

        ws.setContent(range, content);
        Util.printMatrix(ws.getContent(range));
        assertArrayEquals(expectedContent, ws.getContent(range));
    }

    @Test
    public void shouldGetUsedRange() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);

        String range = "A2:B5";
        try {
            Object[][] content = new Object[][]{
                    {"A22", 123},
                    {54.6, 23.5f},
                    {"äöüß", "#?-"},
                    {new SimpleDateFormat("dd.MM.yyyy").parse("03.03.2017"), "=Summe(A3;B3)"}
            };

            ws.setContent(range, content);
            Util.printMatrix(ws.getContent());

            Object[][] actual = ws.getContent(range);
            assertEquals(content[0][0], actual[0][0]);
            assertEquals(content[0][1], ((Double)actual[0][1]).intValue());
            assertEquals(content[1][0], actual[1][0]);
            assertEquals(content[1][1], ((Double)actual[1][1]).floatValue());
            assertEquals(content[2][0], actual[2][0]);
            assertEquals(content[2][1], actual[2][1]);
            assertEquals(content[3][0], actual[3][0]);
            assertEquals(78.1, (Double)actual[3][1], 0.1);
        } catch (ParseException e) {
            fail(e.getClass() + ": " + e.getMessage());
        }
    }

    @Test
    public void shouldGetUnaryContent() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);

        String range = "A1";
        Object content = "test123";

        ws.setContent(range, content);
        assertEquals(content, ws.getContent()[0][0].toString());
    }

    @Test(expected = NullPointerException.class)
    public void shouldSetFillColor() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);

        String range = "A1:B2";
        String range2 = "A1";
        ExcelColor color = ExcelColor.LIGHT_GREEN;
        ExcelColor color2 = ExcelColor.AQUA;

        assertEquals(ExcelColor.XL_NONE, ws.getFillColor(range));
        ws.setFillColor(range, color);
        assertEquals(color, ws.getFillColor(range));

        ws.setFillColor(range2, color2);
        System.out.println(ws.getFillColor(range));
    }

    @Test
    public void
    shouldSetFontColor() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);

        String range = "A1:B2";
        String range2 = "A1";
        ExcelColor color = ExcelColor.LIGHT_GREEN;
        ExcelColor color2 = ExcelColor.AQUA;

        assertEquals(ExcelColor.BLACK, ws.getFontColor(range));
        ws.setFontColor(range, color);
        assertEquals(color, ws.getFontColor(range));

        ws.setFontColor(range2, color2);
        assertEquals(color2, ws.getFontColor(range2));
    }

    @Test
    public void shouldSetBorderColor() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);

        String range = "A1:B2";
        String range2 = "A1";
        ExcelColor color = ExcelColor.LIGHT_GREEN;
        ExcelColor color2 = ExcelColor.AQUA;

        assertEquals(ExcelColor.XL_NONE, ws.getBorderColor(range));
        ws.setBorderColor(range, color);
        assertEquals(color, ws.getBorderColor(range));

        ws.setBorderColor(range2, color2);
    }

    @Test(expected = IllegalArgumentException.class)
    public void shouldSetComment() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);

        String range = "A1";
        String comment = "test comment äöü";
        ws.setComment(range, comment);
        assertEquals(comment, ws.getComment(range));

        ws.setComment("A2:B2", "test123");
    }
}

