package ro.andonescu.excelcomparator;

import org.apache.commons.math.util.MathUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import ro.andonescu.excelcomparator.util.Constants;
import ro.andonescu.excelcomparator.util.XLSUtil;

import java.io.*;
import java.math.BigDecimal;
import java.util.Date;
import java.util.Iterator;

/**
 * Created by iandonescu on 1/9/14.
 */
public class Comparator {
    private String fileOne;
    private String fileTwo;
    private StringBuffer log = new StringBuffer();
    private HSSFWorkbook firstWorkbook;
    private Date compareDate = new Date();

    private int verificationRow = -1;

    public Comparator(String fileOne, String fileTwo) {
        this.fileOne = fileOne;
        this.fileTwo = fileTwo;
    }

    public void compare() {
        try {


            InputStream input = new BufferedInputStream(new FileInputStream(fileOne));
            POIFSFileSystem fs = new POIFSFileSystem(input);
            firstWorkbook = new HSSFWorkbook(fs);
            HSSFSheet sheet = firstWorkbook.getSheetAt(0);
            Iterator rows = sheet.rowIterator();

            if (!XLSUtil.isXLSFile(fileTwo)) {
                fileTwo = new CSVtoXlsTransformer().transformer(fileTwo);
            }
            InputStream input2 = new BufferedInputStream(
                    new FileInputStream(fileTwo));
            POIFSFileSystem fs2 = new POIFSFileSystem(input2);
            HSSFWorkbook wb2 = new HSSFWorkbook(fs2);
            HSSFSheet sheet2 = wb2.getSheetAt(0);


            HSSFCellStyle cs1 = firstWorkbook.createCellStyle();
            cs1.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
            cs1.setFillPattern(CellStyle.SOLID_FOREGROUND);

            HSSFFont f = firstWorkbook.createFont();
            f.setColor(IndexedColors.RED.getIndex());
            cs1.setFont(f);

            while (rows.hasNext()) {

                //iterating each row in the first excel
                verificationRow++;
                HSSFRow row = (HSSFRow) rows.next();
                Iterator celIterator = row.cellIterator();
                HSSFRow row2 = sheet2.getRow(verificationRow);



                for (int j = 0; j  < row.getLastCellNum(); j++) {
                    HSSFCell cellOne = row.getCell(j);
                    // now we will compare the current cel with the one from the other file
                    HSSFCell cellTwo = row2.getCell(j);
                    String result = compareCells(cellOne, cellTwo);
                    if (!result.isEmpty()) {
                        // so we have an error here - log this error in the output file
                        log.append(String.format("row %d - col - %d   --  %s   \n\r ------------------------------------- \n\r", verificationRow, j, result));
                        if (cellOne == null) {
                            cellOne = row.createCell(j);
                        }
                        cellOne.setCellStyle(cs1);
                    }
                }


            }


        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

    public void writeOutputXls() throws IOException {
        if (firstWorkbook == null) {
            throw new RuntimeException("no workbook processed!");
        }

        verifyFolder(Constants.OUTPUT_PATH);
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(Constants.OUTPUT_PATH + String.format("/%s-A-to-B.xls", compareDate.toString().replaceAll("[: ]", "_")));
        firstWorkbook.write(fileOut);
        fileOut.close();
    }

    private void verifyFolder(String filePath) {
        File folder = new File(filePath);
        if (!folder.exists()) {
            folder.mkdirs();
        }
    }

    public void logFile() throws IOException {
        verifyFolder(Constants.OUTPUT_PATH_LOG);

        BufferedWriter out = new BufferedWriter(new FileWriter(Constants.OUTPUT_PATH_LOG + String.format("/%s-log-A-to-B.txt", compareDate.toString().replaceAll("[: ]", "_"))));
        out.write(log.toString());
        out.flush();
        out.close();
    }

    private String compareCells(HSSFCell a, HSSFCell b) {
        StringBuffer sb = new StringBuffer();

        if (a == null && b == null) {
            return "";
        }
        if (a != null && b == null || a == null && b != null) {
//                || a.getCellType() != b.getCellType()) {
            return " different cell types - please check ";
        }
        String valueA;
        boolean isTrue;
        switch ((a.getCellType())) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                if (!XLSUtil.isNumeric(b.getStringCellValue()) || !new Float(a.getNumericCellValue()).equals(new Float(b.getStringCellValue()))) {
                    sb.append(" different values " + a.getNumericCellValue() + " ::: " + b.getStringCellValue());
                }

                break;
            default:
                if (!a.getStringCellValue().equals(b.getStringCellValue())) {
                    sb.append(" different values " + a.getStringCellValue() + " ::: " + b.getStringCellValue());
                }
        }


        return sb.toString();
    }
}
