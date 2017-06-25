package excelcom.test;

import excelcom.api.ExcelColor;
import excelcom.api.ExcelConnection;
import excelcom.api.Workbook;
import excelcom.api.Worksheet;
import excelcom.util.Util;
import org.junit.*;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.net.URISyntaxException;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import static org.junit.Assert.*;

/**
 * Unit tests for Worksheet
 */
public class WorksheetTest {

    private static ExcelConnection connection = null;
    private static Workbook workbook = null;
    private Worksheet worksheet = null;

    @Rule
    public TemporaryFolder tempDir = new TemporaryFolder();

    @BeforeClass
    public static void establishConnection() throws URISyntaxException, InterruptedException {
        connection = ExcelConnection.connect();
        connection.setDisplayAlerts(false);
        workbook = connection.openWorkbook(new File(WorksheetTest.class.getResource("../../test.xlsx").toURI()));
    }

    @AfterClass
    public static void closeConnections() {
        if(workbook != null) {
            workbook.close(false);
        }
        if(connection != null) {
            connection.quit();
        }
    }
    
    @Before
    public void createWorksheet() {
        worksheet = workbook.addWorksheet("test");
    }

    @After
    public void deleteWorksheet() {
        worksheet.delete();
        worksheet = null;
    }

    @Test
    public void shouldModifyWorksheetContent() {
        String range = "A2:B5";
        try {
            Object[][] content = new Object[][]{
                    {"A22", 123},
                    {54.6, 23.5f},
                    {"äöüß", "#?-"},
                    {new SimpleDateFormat("dd.MM.yyyy").parse("03.03.2017"), "=Summe(A3;B3)"}
            };

            worksheet.setContent(range, content);
            //Util.printMatrix(worksheet.getContent(range));

            Object[][] actual = worksheet.getContent(range);
            assertEquals(content[0][0], actual[0][0]);
            assertEquals(content[0][1], ((Double)actual[0][1]).intValue());
            assertEquals(content[1][0], actual[1][0]);
            assertEquals(content[1][1], ((Double)actual[1][1]).floatValue());
            assertEquals(content[2][0], actual[2][0]);
            assertEquals(content[2][1], actual[2][1]);
            assertEquals(content[3][0], actual[3][0]);
            assertEquals(78.1, (Double)actual[3][1], 0.1);

            worksheet.setContent(1,2,4,3, content);
            actual = worksheet.getContent(1,2,4,3);
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
        String range = "A2:B5";
        Object content = 123.5;
        Object[][] expectedContent = new Object[][]{
                {123.5, 123.5},
                {123.5, 123.5},
                {123.5, 123.5},
                {123.5, 123.5}
        };

        worksheet.setContent(range, content);
        //Util.printMatrix(worksheet.getContent(range));
        assertArrayEquals(expectedContent, worksheet.getContent(range));

        worksheet.setContent(1,2,4,3,content);
        assertArrayEquals(expectedContent, worksheet.getContent(1,2,4,3));
    }

    @Test
    public void shouldSetAndGetUnaryRangeContent() {
        worksheet.setUnaryContent("A3", "test");
        worksheet.setUnaryContent("A4", 123);
        worksheet.setUnaryContent("A5", 123.5);
        worksheet.setUnaryContent("A6", "äöüß");
        worksheet.setUnaryContent(new int[]{0,6}, "test123");
        worksheet.setUnaryContent(0,7, "test321");

        assertEquals("test", worksheet.getUnaryContent("A3"));
        assertEquals(123, ((Double)worksheet.getUnaryContent("A4")).intValue());
        assertEquals(123.5, worksheet.getUnaryContent("A5"));
        assertEquals("äöüß", worksheet.getUnaryContent("A6"));
        assertEquals("test123", worksheet.getUnaryContent(new int[]{0,6}));
        assertEquals("test321", worksheet.getUnaryContent(0, 7));
    }

    @Test
    public void shouldGetUsedRange() {
        String range = "A2:B5";
        try {
            Object[][] content = new Object[][]{
                    {"A22", 123},
                    {54.6, 23.5f},
                    {"äöüß", "#?-"},
                    {new SimpleDateFormat("dd.MM.yyyy").parse("03.03.2017"), "=Summe(A3;B3)"}
            };

            worksheet.setContent(range, content);
            //Util.printMatrix(worksheet.getContent());

            Object[][] actual = worksheet.getContent();
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
        String range = "A1";
        Object content = "test123";

        worksheet.setContent(range, content);
        assertEquals(content, worksheet.getContent()[0][0].toString());
    }

    @Test(expected = NullPointerException.class)
    public void shouldSetFillColor() {
        String range = "A1:B2";
        String range2 = "A1";
        ExcelColor color = ExcelColor.LIGHT_GREEN;
        ExcelColor color2 = ExcelColor.AQUA;

        assertEquals(ExcelColor.XL_NONE, worksheet.getFillColor(range));
        worksheet.setFillColor(range, color);
        assertEquals(color, worksheet.getFillColor(range));

        worksheet.setFillColor(range2, color2);
        System.out.println(worksheet.getFillColor(range));
    }

    @Test
    public void
    shouldSetFontColor() {
        String range = "A1:B2";
        String range2 = "A1";
        ExcelColor color = ExcelColor.LIGHT_GREEN;
        ExcelColor color2 = ExcelColor.AQUA;

        assertEquals(ExcelColor.BLACK, worksheet.getFontColor(range));
        worksheet.setFontColor(range, color);
        assertEquals(color, worksheet.getFontColor(range));

        worksheet.setFontColor(range2, color2);
        assertEquals(color2, worksheet.getFontColor(range2));
    }

    @Test
    public void shouldSetBorderColor() {
        String range = "A1:B2";
        String range2 = "A1";
        ExcelColor color = ExcelColor.LIGHT_GREEN;
        ExcelColor color2 = ExcelColor.AQUA;

        assertEquals(ExcelColor.XL_NONE, worksheet.getBorderColor(range));
        worksheet.setBorderColor(range, color);
        assertEquals(color, worksheet.getBorderColor(range));

        worksheet.setBorderColor(range2, color2);
    }

    @Test(expected = IllegalArgumentException.class)
    public void shouldSetComment() {
        String range = "A1";
        String comment = "test comment äöü";
        worksheet.setComment(range, comment);
        assertEquals(comment, worksheet.getComment(range));

        worksheet.setComment("A2:B2", "test123");
    }

}
