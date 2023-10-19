package ApachePOI.JavaClasses;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

public class _02_ApachePOICase_ExecutingData {
    public static void main(String[] args) throws IOException {
        Scanner scanner = new Scanner(System.in);

        System.out.println("Choose the information");
        System.out.println("Username\n" +
                "Password\n" +
                "Address\n" +
                "Zipcode\n" +
                "City\n" +
                "State");
        String userResponse = scanner.nextLine();

        System.out.println(getResult(userResponse));
    }

    public static String getResult(String userResponse) throws IOException {
        String path = "src/test/java/ApachePOI/resources/_02_ApachePOICase_ExecutingData.xlsx";

        FileInputStream fileInputStream = new FileInputStream(path);
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet sheet = workbook.getSheet("Login");

        String returnString = "";
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            if (sheet.getRow(i).getCell(0).toString().equalsIgnoreCase(userResponse)) {
                for (int j = 1; j < sheet.getRow(i).getPhysicalNumberOfCells(); j++) {
                    returnString += sheet.getRow(i).getCell(j);
                }
            }
        }
        return returnString;

    }
}
