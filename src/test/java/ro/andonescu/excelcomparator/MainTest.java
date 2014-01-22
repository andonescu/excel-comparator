package ro.andonescu.excelcomparator;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * Created by iandonescu on 1/14/14.
 */
public class MainTest {
    @Before
    public void setUp() throws Exception {

    }

    @After
    public void tearDown() throws Exception {

    }

    @Test
    public void testMain() throws Exception {
        Main.main(new String[]{"data/expected.xls", "data/actual.xls"});
    }
}
