
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author ppolo
 */
public class DemoWriteFileExcel {

    public static void main(String[] args) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("book");//Tạo Sheet tên là book

        String[][] bookData = {
            {"Name", "Author"},
            {"Head First Java", "Kathy Serria"},
            {"Effective Java", "Joshua Bloch"},
            {"Clean Code", "Robert martin"},
            {"Thinking in Java", "Bruce Eckel"}};

        int rowCount = 0;

        for (String[] aBook : bookData) {
            Row row = sheet.createRow(rowCount++);

            int columnCount = 0;

            for (String field : aBook) {
                Cell cell = row.createCell(columnCount++);

                cell.setCellValue(field);

            }

        }

        try ( FileOutputStream outputStream = new FileOutputStream("DemoWriteExcel.xlsx")) {// Ghi file
            workbook.write(outputStream);
            System.out.println("Ghi thành công");
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

}
