package ro.andonescu.excelcomparator;


import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
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


            HSSFCellStyle textStyle = firstWorkbook.createCellStyle();
            textStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
            textStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

            HSSFFont f = firstWorkbook.createFont();
            f.setColor(IndexedColors.RED.getIndex());
            f.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            textStyle.setFont(f);

            HSSFCellStyle dateStyle = firstWorkbook.createCellStyle();
            dateStyle.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
            dateStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            dateStyle.setFont(f);


            while (rows.hasNext()) {

                //iterating each row in the first excel
                verificationRow++;
                HSSFRow row = (HSSFRow) rows.next();
                HSSFRow row2 = sheet2.getRow(verificationRow);


                for (int j = 0; j < row.getLastCellNum(); j++) {
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
                        if (cellOne.getCellStyle() != null && cellOne.getCellStyle().getDataFormatString() != null) {
                            dateStyle.setDataFormat(cellOne.getCellStyle().getDataFormat());
                            cellOne.setCellStyle(dateStyle);
                        } else {
                            cellOne.setCellStyle(textStyle);
                        }
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
        FileOutputStream fileOut = new FileOutputStream(Constants.OUTPUT_PATH + String.format("/A_B_%s.xls", compareDate.toString().replaceAll("[: ]", "_")));
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

        BufferedWriter out = new BufferedWriter(new FileWriter(Constants.OUTPUT_PATH_LOG + String.format("/A_B_%s.txt", compareDate.toString().replaceAll("[: ]", "_"))));
        out.write(log.toString());
        out.flush();
        out.close();
    }

    private String compareCells(HSSFCell a, HSSFCell b) throws ParseException {
        StringBuffer sb = new StringBuffer();
        if (isBlank(a) && isBlank(b)) {
            return sb.toString();
        }
        if (a != null && b == null || a == null && b != null) {
            return " different cell types - please check ";
        }

        switch ((a.getCellType())) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(a)) {
                    Date aDate = a.getDateCellValue();
                    Date bDate = XLSUtil.toDate(b.getStringCellValue());

                    if (!aDate.equals(bDate)) {
                        sb.append(" different values " + a.getDateCellValue() + " ::: " + b.getStringCellValue());
                    }

                } else if (!XLSUtil.isNumeric(b.getStringCellValue()) || !new Float(a.getNumericCellValue()).equals(new Float(b.getStringCellValue()))) {
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
