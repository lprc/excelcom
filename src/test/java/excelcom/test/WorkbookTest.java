package excelcom.test;

import excelcom.api.ExcelConnection;
import excelcom.api.Workbook;
import excelcom.api.Worksheet;
import org.junit.*;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

/**
 * Unit tests for Workbook
 */
public class WorkbookTest {

    private static ExcelConnection connection = null;
    private static Workbook workbook = null;

    @Rule
    public TemporaryFolder tempDir = new TemporaryFolder();

    @BeforeClass
    public static void establishConnection() throws URISyntaxException, InterruptedException {
        connection = ExcelConnection.connect();
        connection.setDisplayAlerts(false);
        workbook = connection.openWorkbook(new File(WorkbookTest.class.getResource("../../test.xlsx").toURI()));
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

    @Test
    public void shouldAddWorksheet() {
        Worksheet ws = workbook.addWorksheet("test");
        assertNotNull(ws);
        ws.delete();
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
    public void shouldSaveWorkbookAs() throws IOException {
        File f = tempDir.newFile("test123.xlsx");
        workbook.saveAs(f);
        assertTrue(f.exists());
    }

}
