package excelcom.test;

import excelcom.api.ExcelConnection;
import excelcom.api.Workbook;
import org.junit.*;

import java.io.File;
import java.net.URISyntaxException;

import static org.junit.Assert.*;

/**
 * Unit tests for ExcelConnection
 */
public class ExcelConnectionTest {

    private static ExcelConnection connection = null;

    @BeforeClass
    public static void establishConnection() throws URISyntaxException, InterruptedException {
        connection = ExcelConnection.connect();
        connection.setDisplayAlerts(false);
    }

    @AfterClass
    public static void closeConnections() {
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
    public void shouldOpenWorkbook() throws URISyntaxException {
        assertNotNull(connection.openWorkbook(new File(getClass().getResource("../../test.xlsx").toURI())));
    }

    @Test
    public void shouldUseActiveInstance() {
        ExcelConnection con2 = ExcelConnection.connect(true);
        assertNotNull(con2);
        con2.quit();
        assertNotNull(connection.getVersion());
    }

    @Test
    public void shouldCreateWorkbook() {
        File f = new File(System.getProperty("java.io.tmpdir") + "/test123.xlsx");
        Workbook wb = connection.newWorkbook(f);
        assertNotNull(wb);
        wb.close(true);
        assertTrue(f.exists());
        assertTrue(f.delete());
    }
}
