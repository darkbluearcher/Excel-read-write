import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
    public class excelwrite {

        private static final String FILE_NAME = "C:\\Users\\batuh\\Desktop\\deneme1.xlsx";

        public static void main(String[] args) {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Çalışma Kağıdı");
            Object[][] datatypes = {
                    {"Ad", "Soyad", "Sınıf"},
                    {"Batuhan", "Kahraman", "A12/A"},
                    {"float", "Bilici", "A12/A"},
                    {"Berna", "Gürek", "A12/A"},
                    {"deneme", "deneme", "deneme"},

            };

            int rowNum = 0;
            System.out.println("Creating excel");

            for (Object[] datatype : datatypes) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (Object field : datatype) {
                    Cell cell = row.createCell(colNum++);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }
            }

            try {
                FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
                workbook.write(outputStream);
                workbook.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }

            System.out.println("Başarılı");
        }
    }

