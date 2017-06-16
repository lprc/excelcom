package excelcom.test;

import excelcom.api.ExcelConnection;
import org.junit.*;

import java.io.File;
import java.net.URISyntaxException;

import static org.junit.Assert.assertNotNull;

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
}
