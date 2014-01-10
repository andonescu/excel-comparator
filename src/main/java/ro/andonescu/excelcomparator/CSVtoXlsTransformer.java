package ro.andonescu.excelcomparator;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
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

    public String transformer(String fName, String refXLS) throws IOException {

        createOutputFolder();

        InputStream input = new BufferedInputStream(
                new FileInputStream(refXLS));
        POIFSFileSystem fs = new POIFSFileSystem(input);
        HSSFWorkbook  firstWorkbook = new HSSFWorkbook(fs);
        HSSFSheet refSHEET = firstWorkbook.getSheetAt(0);

        ArrayList arList = null;
        ArrayList al = null;
        String thisLine;
        int count = 0;
        File file = new File(fName);
        FileInputStream fis = new FileInputStream(file);
        DataInputStream myInput = new DataInputStream(fis);

        int i = 0;
        arList = new ArrayList();
        while ((thisLine = myInput.readLine()) != null) {
            al = new ArrayList();
            String strar[] = thisLine.split("\t");
            for (int j = 0; j < strar.length; j++) {
                al.add(strar[j]);
            }
            arList.add(al);
            System.out.println();
            i++;
        }

        try {
            HSSFWorkbook hwb = new HSSFWorkbook();
            HSSFSheet sheet = hwb.createSheet("new sheet");
            HSSFCellStyle style = hwb.createCellStyle();


            for (int k = 0; k < arList.size(); k++) {
                ArrayList rowDataList = (ArrayList) arList.get(k);
                HSSFRow row = sheet.createRow((short) 0 + k);
                HSSFRow refRow =    refSHEET.getRow(k);
                for (int p = 0; p < rowDataList.size(); p++) {
                    HSSFCell cell = row.createCell( p);

                    String columnData = rowDataList.get(p).toString().trim().replace("\"","");
//                    if (XLSUtil.isNumeric(columnData)) {
//
////                        if ( refRow.getCell(p) != null && refRow.getCell(p).getCellStyle() != null) {
////                            style.setDataFormat(refRow.getCell(p).getCellStyle().getDataFormat());
////                        }
//                        cell.setCellStyle(style);
//                        cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
//
//                        cell.setCellValue(columnData);
//
//                    } else {
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        cell.setCellValue(columnData);
//                    }


                }
                System.out.println();
            }


            String newFilePath = String.format("%s/%s-%s.xls", Constants.TEMP_FOLDER,
                    new Date().toString().replaceAll("[ :]", "_"),
                   file.getName());
            FileOutputStream fileOut = new FileOutputStream(newFilePath);
            hwb.write(fileOut);
            fileOut.close();
            System.out.println("Your excel file has been generated");
            myInput.close();

            return newFilePath;
        } catch (Exception ex) {
            ex.printStackTrace();
        } //main method ends

        throw new RuntimeException(" no gen possible!");
    }

    private void createOutputFolder() {
        XLSUtil.verifyAndCreateFolder(Constants.OUTPUT_PATH);
        XLSUtil.verifyAndCreateFolder(Constants.TEMP_FOLDER);
    }

}
