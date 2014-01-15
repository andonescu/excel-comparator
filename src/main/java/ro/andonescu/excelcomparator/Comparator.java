package ro.andonescu.excelcomparator;


import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.IndexedColors;
import ro.andonescu.excelcomparator.util.Constants;
import ro.andonescu.excelcomparator.util.XLSUtil;

import java.io.*;
import java.text.ParseException;
import java.util.Date;
import java.util.Iterator;

/**
 * Created by iandonescu on 1/9/14.
 */
public class Comparator {
    private String expectedFile;
    private String actualFile;
    private StringBuffer log = new StringBuffer();
    private HSSFWorkbook expectedWorkbook;
    private Date compareDate = new Date();

    private int verificationRow = -1;

    public Comparator(String expectedFilePath, String actualFilePath) {
        this.expectedFile = expectedFilePath;
        this.actualFile = actualFilePath;
    }

    /**
     * Compare the files
     */
    public void compare() {
        try {


            HSSFSheet expectedSheet = getExpectedWorkbookFirstSheet(expectedFile);
            Iterator rows = expectedSheet.rowIterator();

            if (!XLSUtil.isXLSFile(actualFile)) {
                actualFile = new CSVtoXlsTransformer().transformer(actualFile, expectedSheet);
            }

            HSSFSheet actualSheet = getActualWorkbookFirstSheet(actualFile);

            doActualComparasion(rows, actualSheet);

        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }

    private HSSFSheet getExpectedWorkbookFirstSheet(String file) throws IOException {
        expectedWorkbook = getWorkbook(file);
        return expectedWorkbook.getSheetAt(0);
    }

    private HSSFWorkbook getWorkbook(String file) throws IOException {
        InputStream input = new BufferedInputStream(new FileInputStream(file));
        POIFSFileSystem fs = new POIFSFileSystem(input);
        return new HSSFWorkbook(fs);
    }

    private HSSFSheet getActualWorkbookFirstSheet(String file) throws IOException {
        HSSFWorkbook actualWorkbook = getWorkbook(file);
        return actualWorkbook.getSheetAt(0);
    }

    private void doActualComparasion(Iterator rows, HSSFSheet sheet2) throws ParseException {
        HSSFCellStyle textStyle = null;
        HSSFCellStyle dateStyle = null;

        HSSFFont redFont = getRedFont();

        while (rows.hasNext()) {

            //iterating each row in the first excel
            verificationRow++;
            HSSFRow row = (HSSFRow) rows.next();
            HSSFRow row2 = sheet2.getRow(verificationRow);


            for (int j = 0; j < row.getLastCellNum(); j++) {
                // now we will compare the current cel with the one from the other file
                HSSFCell cellOne = row.getCell(j);
                HSSFCell cellTwo = row2.getCell(j);

                String result = verifyCellsValues(cellOne, cellTwo);

                if (!result.isEmpty()) {
                    // so we have an error here - log this error in the output file
                    log.append(String.format("row %d - col - %d   --  %s   \n\r ------------------------------------- \n\r", verificationRow, j, result));
                    boolean hasTextStyle = true;
                    if (cellOne == null) {
                        cellOne = row.createCell(j);
                    }

                    if (cellOne.getCellStyle() != null
                            && cellOne.getCellStyle().getDataFormatString() != null
                            && cellOne.getCellType() == HSSFCell.CELL_TYPE_NUMERIC
                            && HSSFDateUtil.isCellDateFormatted(cellOne)) {

                        dateStyle = updateStyleIfNeeded(dateStyle, redFont, cellOne);
                        cellOne.setCellStyle(dateStyle);
                        hasTextStyle = false;
                    }

                    if (hasTextStyle) {
                        textStyle = updateStyleIfNeeded(textStyle, redFont, cellOne);
                        cellOne.setCellStyle(textStyle);
                    }
                }
            }


        }
    }

    private HSSFCellStyle updateStyleIfNeeded(HSSFCellStyle dateStyle, HSSFFont redFont, HSSFCell cellOne) {
        if (dateStyle == null) {
            dateStyle = updateStyle(redFont, cellOne);
        }
        return dateStyle;
    }

    private HSSFCellStyle updateStyle(HSSFFont redFont, HSSFCell cell) {
        HSSFCellStyle textStyle;
        textStyle = expectedWorkbook.createCellStyle();
        textStyle.cloneStyleFrom(cell.getCellStyle());
        textStyle.setFont(redFont);
        return textStyle;
    }

    private HSSFFont getRedFont() {
        HSSFFont redFont = expectedWorkbook.createFont();
        redFont.setColor(IndexedColors.RED.getIndex());
        redFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        return redFont;
    }

    public void writeOutputXls() throws IOException {
        if (expectedWorkbook == null) {
            throw new RuntimeException("no workbook processed!");
        }

        verifyFolder(Constants.OUTPUT_PATH);
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(Constants.OUTPUT_PATH + String.format("/A_B_%s.xls", compareDate.toString().replaceAll("[: ]", "_")));
        expectedWorkbook.write(fileOut);
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

        BufferedWriter out = new BufferedWriter(new FileWriter(Constants.OUTPUT_PATH_LOG + String.format("/A_B_%s.txt", compareDate.toString().replaceAll("[: ]", "_"))));
        out.write(log.toString());
        out.flush();
        out.close();
    }

    private String verifyCellsValues(HSSFCell a, HSSFCell b) throws ParseException {
        StringBuffer sb = new StringBuffer();
        if (isBlank(a) && isBlank(b)) {
            return sb.toString();
        }
        if (a != null && b == null || a == null && b != null) {
            return " different cell types - please check ";
        }
//
        switch ((a.getCellType())) {
            case HSSFCell.CELL_TYPE_NUMERIC:

                if (HSSFDateUtil.isCellDateFormatted(a)) {
                    Date aDate = a.getDateCellValue();
                    Date bDate = b.getDateCellValue();

                    if (!aDate.equals(bDate)) {
                        sb.append(" different values " + a.getDateCellValue() + " ::: " + b.getDateCellValue());
                    }

                } else {

                    if (a.getNumericCellValue() != b.getNumericCellValue()) {
                        sb.append(" different values " + a.getNumericCellValue() + " ::: " + b.getNumericCellValue());
                    }
                }

                break;
            default:
                if (!a.getStringCellValue().equals(b.getStringCellValue())) {
                    sb.append(" different values " + a.getStringCellValue() + " ::: " + b.getStringCellValue());
                }
        }


        return sb.toString();
    }

    private boolean isBlank(HSSFCell cell) {
        if (cell == null) {
            return true;
        }
        switch ((cell.getCellType())) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                return false;
            default:
                return StringUtils.isBlank(cell.getStringCellValue());
        }
    }
}
