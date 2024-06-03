package org.example;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;

@AllArgsConstructor
@NoArgsConstructor
@Data
public class Manager {
    private String email;
    private String password;
    private String actions;

    public static Manager login(User user) {
        try {
            String filePath = "F:/OnlineStoreAppBackendAPI/data/OnlineStoreAppDatabase.xlsx";
            Workbook workbook;
            FileInputStream fis = new FileInputStream(filePath);
            workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheet("Managers");
            if (sheet != null) {
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);

                    Cell cell = row.getCell(0);
                    String email = cell.getStringCellValue();
                    cell = row.getCell(1);
                    String password;
                    if (cell.getCellType() == CellType.STRING) {
                        password = cell.getStringCellValue();
                    } else {
                        password = String.valueOf((int)(cell.getNumericCellValue()));
                    }
                    cell = row.getCell(2);
                    String actions = cell.getStringCellValue();

                    if (user.getEmail().equals(email) && user.getPassword().equals(password)) {
                        Manager manager = new Manager(email, password, actions);
                        return manager;
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Exception occurred while manager login");
            e.printStackTrace();
        }
        return null;
    }

    public static String fetchManagerActions(String reqEmail) {
        try {
            String filePath = "F:/OnlineStoreAppBackendAPI/data/OnlineStoreAppDatabase.xlsx";
            Workbook workbook;
            FileInputStream fis = new FileInputStream(filePath);
            workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheet("Managers");
            if (sheet != null) {
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);

                    Cell cell = row.getCell(0);
                    String email = cell.getStringCellValue();
                    cell = row.getCell(2);
                    String actions = cell.getStringCellValue();

                    if (reqEmail.equals(email)) {
                        return actions;
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Exception occurred while fetching manager actions");
            e.printStackTrace();
        }
        return null;
    }
}
