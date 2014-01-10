package ro.andonescu.excelcomparator;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import ro.andonescu.excelcomparator.util.Constants;

import java.io.*;
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

    public HSSFCell getNextCell(HSSFRow row) {
        Iterator cells = row.cellIterator();
        return (HSSFCell) cells.next();

    }

    public void compare() {
        try {



            InputStream input = new BufferedInputStream(
                    new FileInputStream(fileOne));
            POIFSFileSystem fs = new POIFSFileSystem(input);
            firstWorkbook = new HSSFWorkbook(fs);
            HSSFSheet sheet = firstWorkbook.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
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
                boolean flag = false;
                //iterating each row in the first excel
                verificationRow++;
                HSSFRow row = (HSSFRow) rows.next();

                Iterator celIterator = row.cellIterator();
                int verificationColumn = -1;
                while (celIterator.hasNext()) {
                    verificationColumn++;
                    HSSFCell cellOne = (HSSFCell) celIterator.next();
                    // now we will compare the current cel with the one from the other file

                    HSSFRow row2 = sheet2.getRow(verificationRow);
                    HSSFCell cellTwo = row2.getCell(verificationColumn);
                    String result = compareCells(cellOne, cellTwo);
                    if (!result.isEmpty()) {
                        // so we have an error here - log this error in the output file
                        log.append(String.format("row %d - col - %d   --  %s   \n\r ------------------------------------- \n\r", verificationRow, verificationColumn, result));
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

        verifyFolder();
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(Constants.OUTPUT_PATH + String.format("/%s-A-to-B.xls", compareDate.toString().replaceAll("[: ]", "_")));
        firstWorkbook.write(fileOut);
        fileOut.close();
    }

    private void verifyFolder() {
        File folder = new File(Constants.OUTPUT_PATH);
        if (!folder.exists()) {
            folder.mkdir();
        }
    }

    public void logFile() throws IOException {
        verifyFolder();

        BufferedWriter out = new BufferedWriter(new FileWriter(Constants.OUTPUT_PATH + String.format("/%s-log-A-to-B.txt", compareDate.toString().replaceAll("[: ]", "_"))));
        out.write(log.toString());
        out.flush();
        out.close();
    }

    private String compareCells(HSSFCell a, HSSFCell b) {
        StringBuffer sb = new StringBuffer();

        if (a == null && b == null) {
            return "";
        }
        if (a != null && b == null || a== null && b != null || a.getCellType() != b.getCellType()) {
            return " different cell types - please check ";
        }

        switch ((a.getCellType())) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                if (a.getNumericCellValue() != b.getNumericCellValue()) {
                    sb.append(" different values " + a.getNumericCellValue() + " ::: " + b.getNumericCellValue());
                }
                break;
            case HSSFCell.CELL_TYPE_STRING:
                if (!a.getStringCellValue().equals(b.getStringCellValue())) {
                    sb.append(" different values " + a.getStringCellValue() + " ::: " + b.getStringCellValue());
                }
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                if (!a.getStringCellValue().equals(b.getStringCellValue())) {
                    sb.append(" different values " + a.getStringCellValue() + " ::: " + b.getStringCellValue());
                }
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                if (a.getBooleanCellValue() != b.getBooleanCellValue()) {
                    sb.append(" different values " + a.getBooleanCellValue() + " ::: " + b.getBooleanCellValue());
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
