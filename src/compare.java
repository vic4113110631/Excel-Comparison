import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class compare {
    public enum  EDU_Field{
        SOURCEID,
        LANGUAGES,
        SCHOOL_NAME,
        CITY,
        COUNTRY,
        MAJOR,
        DEGREE,
        PERIOD_START,
    }
    public enum RP_Field{
        CRISID,
        UUID,
        SOURCEREF,
        SOURCEID
    }
    public static void readFile(String FILE_PATH_1, String FILE_PATH_2){
        try {
            InputStream FILE_1 = new FileInputStream(FILE_PATH_1);
            InputStream FILE_2 = new FileInputStream(FILE_PATH_2);

            HSSFWorkbook RP = new HSSFWorkbook(FILE_1);
            XSSFWorkbook EDU = new XSSFWorkbook(FILE_2);

            HSSFSheet entites = RP.getSheet("main_entities");
            XSSFSheet educations = EDU.getSheet("education");
            Iterator<Row> rows = entites.iterator();

            while (rows.hasNext()) {
                Row row = rows.next();

                Iterator<Cell> cells = row.iterator();
                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    if (cell.getCellTypeEnum() == CellType.STRING) {
                        System.out.print(cell.getStringCellValue() + "--");
                    } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                        System.out.print(cell.getNumericCellValue() + "--");
                    }
                }
                System.out.println();
            }


        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void main(String[] args) {
        readFile("src/AH_RP.xls", "src/AH_Education.xlsx");
    }
}

