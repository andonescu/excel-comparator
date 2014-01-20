package ro.andonescu.excelcomparator;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import ro.andonescu.excelcomparator.util.Constants;
import ro.andonescu.excelcomparator.util.XLSUtil;

import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by iandonescu on 1/10/14.
 */
public class CSVtoXlsTransformer {

    private HSSFWorkbook hwb;
    private Map<String, HSSFCellStyle> styles = new HashMap<String, HSSFCellStyle>();

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
            hwb = new HSSFWorkbook();
            HSSFSheet sheet = hwb.createSheet(compareSheet.getSheetName());


            for (int i = 0; i < arList.size(); i++) {
                ArrayList rowDataList = (ArrayList) arList.get(i);
                HSSFRow row = sheet.createRow((short) 0 + i);
                for (int j = 0; j < rowDataList.size(); j++) {

                    HSSFCell cell = row.createCell(j);
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

            HSSFCellStyle newStyle = getHssfCellStyle(compareCell);
            cell.setCellStyle(newStyle);

            switch ((compareCell.getCellType())) {
                case HSSFCell.CELL_TYPE_NUMERIC:
                    newStyle.setDataFormat(compareCell.getCellStyle().getDataFormat());
                    if (HSSFDateUtil.isCellDateFormatted(compareCell)) {
                        Date bDate = XLSUtil.toDate(columnData);
                        cell.setCellValue(bDate);
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

    private HSSFCellStyle getHssfCellStyle(HSSFCell compareCell) {
        String key = String.format("%d-%s", compareCell.getCellType(),
                compareCell.getCellStyle().getDataFormatString());


        if (styles.containsKey(key)) {
            return styles.get(key);
        }

        HSSFCellStyle newStyle = hwb.createCellStyle();
        newStyle.cloneStyleFrom(compareCell.getCellStyle());
        styles.put(key, newStyle);
        return newStyle;
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
