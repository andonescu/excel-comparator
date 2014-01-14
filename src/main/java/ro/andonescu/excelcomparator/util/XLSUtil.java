package ro.andonescu.excelcomparator.util;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by iandonescu on 1/10/14.
 */
public class XLSUtil {

    public static boolean isXLSFile(String file) {
        boolean isXLS = false;
        try {
            InputStream input = new BufferedInputStream(new FileInputStream(file));
            POIFSFileSystem fs = new POIFSFileSystem(input);
            HSSFWorkbook workbook = new HSSFWorkbook(fs);
            HSSFSheet sheet = workbook.getSheetAt(0);
            sheet.rowIterator();

            // at this moment we know that it is a xls file
            isXLS = true;
            input.close();
        } catch (IOException e) {
//            e.printStackTrace();
        }

        return isXLS;

    }

    public static void verifyAndCreateFolder(String path) {
        File folder = new File(path);
        if (!folder.exists()) {
            folder.mkdir();
        }
    }

    public static boolean isNumeric(String s) {
        return s.matches("[+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)");
    }

    public static Date toDate(String date){
        String[] dateFormats = {"MM/dd/yyyy HH:mm:ss", "MM/dd/yyyy HH:mm"};
        Date finalDate = null;
        for(String format : dateFormats) {
            try {
                finalDate = getDateByFormat(date.trim(), format);
            } catch (ParseException e) {
               // this format is not ok
                //TODO: modify this to use regex
//                e.printStackTrace();
            }
        }

        if (finalDate == null) {
            throw new RuntimeException("unknown date format - please update the code: " + date);
        }

        return finalDate;
    }

    public static Date getDateByFormat(String date, String format) throws ParseException {
        SimpleDateFormat dateFormat = new SimpleDateFormat(format);
        return dateFormat.parse(date);
    }
}
