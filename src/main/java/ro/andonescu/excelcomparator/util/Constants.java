package ro.andonescu.excelcomparator.util;

/**
 * Created by iandonescu on 1/9/14.
 */
public class Constants {

    public static String OUTPUT_PATH = ".";
    public static String OUTPUT_PATH_LOG;
    public static String OUTPUT_PATH_DIFF;
    public static String OUTPUT_PATH_COMPARED;

    public static String TEMP_FOLDER;

    static {
        update();
    }


    public static void update() {
        TEMP_FOLDER = OUTPUT_PATH + "/temp";
        OUTPUT_PATH_LOG = TEMP_FOLDER + "/logs";
        OUTPUT_PATH_DIFF = TEMP_FOLDER + "/differences";
        OUTPUT_PATH_COMPARED = TEMP_FOLDER + "/compared";
    }
}
