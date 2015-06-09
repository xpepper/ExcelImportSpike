import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Spike {
    public static void main(String[] args) {
        if (args.length > 0) {
            String filePath = args[0];
            readFile(new File(filePath));
        } else {
            System.out.println("Missing command line arguments!");
            return;
        }

    }

    private static void readFile(File file) {
        try {
            Workbook workbook = WorkbookFactory.create(file);
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                System.out.println("=== Sheet " + sheet.getSheetName());

                for (Row row : sheet) {
                    System.out.println("Row " + (row.getRowNum() + 1));
                    for (Cell cell : row) {
                        System.out.println("\t '" + cell.toString() + "'");
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Generated exception: " + e);
        }

    }
}
