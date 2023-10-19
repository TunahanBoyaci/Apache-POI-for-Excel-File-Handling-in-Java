package ApachePOI.JavaClasses;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

public class _01_ApachePOICase_ReadingData {
    public static void main(String[] args) throws IOException {
        String path = "src/test/java/ApachePOI/resources/_01_ApachePOICase_ReadingData.xlsx";

        FileInputStream fileInputStream = new FileInputStream(path);
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet sheet = workbook.getSheet("Sheet1");

        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            for (int j = 0; j < sheet.getRow(i).getPhysicalNumberOfCells(); j++) {
                System.out.print(" (Row: " + i + " Cell: " + j + "): " + sheet.getRow(i).getCell(j));
            }
            System.out.println("");
        }


    }
}
