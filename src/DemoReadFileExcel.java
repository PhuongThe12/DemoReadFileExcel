
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author ppolo
 */
public class DemoReadFileExcel {

    public static void main(String[] args) {
        File file = new File("DemoExcel.xlsx");
        readByCell(file, 4, 1);
    }

    public static void readAll(File file) {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(file);//đọc file 

            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);//Lấy trang đầu tiên trong file excel
            Iterator<Row> itr = sheet.iterator();//Lấy ra các dòng

            while (itr.hasNext()) {//Lặp đến khi hết các dòng trong excel
                Row row = itr.next();//Lấy dòng tiếp theo
                Iterator<Cell> cellItr = row.iterator(); // Lấy ra các ô trong dòng row

                while (cellItr.hasNext()) {
                    Cell cell = cellItr.next();
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                            System.out.print(cell.getStringCellValue() + "\t\t\t");
                            break;
                        case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
                            System.out.print(cell.getNumericCellValue() + "\t\t\t");
                            break;
                        default:
                    }
                }

                System.out.println("");

            }

            fis.close();
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    public static void readByCell(File file, int vRow, int vColumn) {
        Workbook wb = null;
        try {

            FileInputStream fis = new FileInputStream(file);

            wb = new XSSFWorkbook(fis);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
        Sheet sheet = wb.getSheetAt(0);   //Lấy trang đầu tiên trong file excel
        Row row = sheet.getRow(vRow); //Lầy dòng thứ vRow
        Cell cell = row.getCell(vColumn); //Lấy ô thứ vColumn trong dòng thứ vRow

        System.out.println(cell.getStringCellValue());//In ra giá trị trong ô

    }

}
