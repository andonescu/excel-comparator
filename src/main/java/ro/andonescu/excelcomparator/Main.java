package ro.andonescu.excelcomparator;

import ro.andonescu.excelcomparator.util.Constants;

import java.io.IOException;

/**
 * Created by iandonescu on 1/9/14.
 */
public class Main {
    public static void main(String[] args) throws IOException {
        verifyArguments(args);

        updateConstants(args);

        compareFiles(args);
    }

    private static void compareFiles(String[] args) throws IOException {
        Comparator comparator = new Comparator(args[0], args[1]);
        comparator.compare();
        comparator.writeOutputXls();
        comparator.logFile();
    }

    private static void verifyArguments(String[] args) {
        if (args.length < 2) {
            throw new RuntimeException("example of usage  $: java -jar comparator.jar d:/A.xls d:/b.xls");
        }
    }

    private static void updateConstants(String[] args) {
        String defaultUser = "loghinc120";
        if (args.length >= 3) {
            defaultUser = args[2];
        }

        Constants.OUTPUT_PATH = String.format(Constants.OUTPUT_PATH, defaultUser);
        Constants.update();
    }
}
