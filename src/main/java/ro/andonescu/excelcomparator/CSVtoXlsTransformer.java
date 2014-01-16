package ro.andonescu.excelcomparator;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import ro.andonescu.excelcomparator.util.Constants;
import ro.andonescu.excelcomparator.util.XLSUtil;

import java.io.*;
import java.util.ArrayList;
import java.util.Date;

/**
 * Created by iandonescu on 1/10/14.
 */
public class CSVtoXlsTransformer {

    private HSSFCellStyle normalStyle ;
    private HSSFCellStyle dateStyle ;
    /**
     * Transforms the given file, in to a xls one, based on another
     *
     * @param filePathToBeTransformed
     * @param compareSheet
     * @return
     * @throws IOException
     */
    public String transformer(String filePathToBeTransformed, HSSFSheet compareSheet) throws IOException {

        createOutputFolder();

        ArrayList arList = null;
        ArrayList al = null;
        String thisLine;
        File file = new File(filePathToBeTransformed);
        FileInputStream fis = new FileInputStream(file);
        DataInputStream myInput = new DataInputStream(fis);

        arList = new ArrayList();
        while ((thisLine = myInput.readLine()) != null) {
            al = new ArrayList();
            String data[] = thisLine.split("\t");
            for (int j = 0; j < data.length; j++) {
                al.add(data[j]);
            }
            arList.add(al);
        }

        try {
            HSSFWorkbook hwb = new HSSFWorkbook();
            HSSFSheet sheet = hwb.createSheet(compareSheet.getSheetName());

            normalStyle = hwb.createCellStyle();
            dateStyle = hwb.createCellStyle();

            for (int i = 0; i < arList.size(); i++) {
                ArrayList rowDataList = (ArrayList) arList.get(i);
                HSSFRow row = sheet.createRow((short) 0 + i);
                for (int j = 0; j < rowDataList.size(); j++) {

                    HSSFCell cell = row.createCell(j);
                    cell.setCellStyle(normalStyle);
                    String columnData = cleanData(rowDataList, j);
                    HSSFCell compareCell = compareSheet.getRow(i).getCell(j);

                    storeDataToCell(cell, columnData, compareCell);

                }
            }


            return writeTheOutput(file, myInput, hwb);

        } catch (Exception ex) {
            ex.printStackTrace();
        } //main method ends

        throw new RuntimeException(" no gen possible!");
    }

    private String writeTheOutput(File file, DataInputStream myInput, HSSFWorkbook hwb) throws IOException {
        File folder = new File(Constants.OUTPUT_PATH_COMPARED);
        folder.mkdirs();

        String newFilePath = String.format("%s/%s_%s.xls", Constants.OUTPUT_PATH_COMPARED,
                file.getName(), new Date().toString().replaceAll("[ :]", "_")
        );

        FileOutputStream fileOut = new FileOutputStream(newFilePath);
        hwb.write(fileOut);
        fileOut.close();
        System.out.println("Your excel file has been generated");
        myInput.close();
        return newFilePath;
    }

    private void storeDataToCell(HSSFCell cell, String columnData, HSSFCell compareCell) {
        if (compareCell != null) {
            cell.setCellType(compareCell.getCellType());

            switch ((compareCell.getCellType())) {
                case HSSFCell.CELL_TYPE_NUMERIC:

                    if (HSSFDateUtil.isCellDateFormatted(compareCell)) {
                        Date bDate = XLSUtil.toDate(columnData);
                        cell.setCellValue(bDate);
                        dateStyle.setDataFormat(compareCell.getCellStyle().getDataFormat());
                        cell.setCellStyle(dateStyle);
                    } else {
                        cell.setCellValue(new Double(columnData));
                    }
                    break;
                default:
                    cell.setCellValue(columnData);
                    break;
            }

        } else {
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(columnData);
        }
    }

    private String cleanData(ArrayList rowDataList, int j) {
        String columnData = rowDataList.get(j).toString().trim();
        if (columnData.startsWith("\"")) {
            columnData = columnData.substring(1, columnData.length() - 1);
        }
        if (columnData.endsWith("\"")) {
            columnData = columnData.substring(0, columnData.length() - 2);
        }
        columnData = columnData.replaceAll("\"\"", "\"");
        return columnData;
    }

    private void createOutputFolder() {
        XLSUtil.verifyAndCreateFolder(Constants.OUTPUT_PATH);
        XLSUtil.verifyAndCreateFolder(Constants.TEMP_FOLDER);
    }

}
