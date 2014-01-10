package ro.andonescu.excelcomparator;

import java.io.IOException;

/**
 * Created by iandonescu on 1/9/14.
 */
public class Main {
    public static void main(String[] args) throws IOException {
        if (args.length != 2) {
            throw new RuntimeException("example of usage  $: java -jar comparator.jar d:/A.xls d:/b.xls");
        }
        Comparator comparator = new Comparator(args[0], args[1]);
        comparator.compare();
        comparator.writeOutputXls();
        comparator.logFile();
    }
}
