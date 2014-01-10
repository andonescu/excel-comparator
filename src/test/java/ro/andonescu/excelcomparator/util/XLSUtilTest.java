package ro.andonescu.excelcomparator.util;


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
    public void testIsNumber() throws Exception {
        Assert.assertEquals(true,XLSUtil.isNumeric("33"));
        Assert.assertEquals(true, XLSUtil.isNumeric("22.43"));
        Assert.assertEquals(true, XLSUtil.isNumeric("123123"));

        Assert.assertEquals(false, XLSUtil.isNumeric("test.csv"));
    }
}
