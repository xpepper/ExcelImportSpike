import java.io.File;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Spike {
    public static void main(String[] args) throws Exception {
        String pathname = "/Users/pierodibello/" + "Downloads/2015.05.21_MonitoraggioOO.PP.xls";
        Workbook workbook = WorkbookFactory.create(new File(pathname));

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
    }
}
