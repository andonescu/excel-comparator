package ro.andonescu.excelcomparator.util;


import org.joda.time.DateTime;
import org.junit.Assert;

/**
 * Created by iandonescu on 1/10/14.
 */
public class XLSUtilTest {
    @org.junit.Before
    public void setUp() throws Exception {

    }

    @org.junit.After
    public void tearDown() throws Exception {

    }

    @org.junit.Test
    public void testIsXLSFile_isOK() throws Exception {
        Assert.assertEquals(true, XLSUtil.isXLSFile("test.xls"));
    }

    @org.junit.Test
    public void testIsXLSFile_isNotOK() throws Exception {
        Assert.assertEquals(false, XLSUtil.isXLSFile("test.csv"));
    }

    @org.junit.Test
       public void toDate_withTimeStamp() throws Exception {
        Assert.assertNotNull(XLSUtil.toDate("2013-12-09 13:31:10.454545"));

        System.out.println(new DateTime(XLSUtil.toDate("2013-12-09 13:31:10.454545").getTime()));
    }

    @org.junit.Test
    public void toDate_NoTimestamp() throws Exception {
        Assert.assertNotNull(XLSUtil.toDate("2013-12-09 13:31:10"));

        System.out.println(new DateTime(XLSUtil.toDate("2013-12-09 13:31:10").getTime()));
    }
}
