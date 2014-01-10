package ro.andonescu.excelcomparator.util;

/**
 * Created by iandonescu on 1/9/14.
 */
public class Constants {

    public static String OUTPUT_PATH = "C:/Users/%s/Desktop/OES-PoC";
    public static String OUTPUT_PATH_LOG;

    public static String TEMP_FOLDER;

    static {
        update();
    }


    public static void update(){
        TEMP_FOLDER = OUTPUT_PATH + "/temp";
        OUTPUT_PATH_LOG  = TEMP_FOLDER + "/log";
    }
}
