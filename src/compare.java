import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import static java.lang.Boolean.FALSE;
import static java.lang.Boolean.TRUE;

public class compare {
    public enum  EDU_Field{
        TED_ID,
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

            Iterator<Row> rows_en = entites.iterator();
            Iterator<Row> rows_edu = educations.iterator();

            // Skip first entry
            if (rows_en.hasNext())
                rows_en.next();
            if (rows_edu.hasNext())
                rows_edu.next();

            // To control Rows education Loop
            Row previous_row = null;
            Boolean isNewLoop = FALSE;

            while (rows_en.hasNext()) {
                Row row_en = rows_en.next();
                // get source ID from AH_RP.xls and compare to education sheet
                // default source ID from AH_RP.xls is string, so convert it to integer
                Integer SOURCEID_EN = Integer.parseInt(row_en.getCell(RP_Field.SOURCEID.ordinal()).getStringCellValue());

                Row row_edu = null;
                while (rows_edu.hasNext()){
                    if(isNewLoop.equals(TRUE)) {
                        row_edu = previous_row;
                        isNewLoop = FALSE;
                    }else{
                        row_edu = rows_edu.next();
                    }
                    // source ID and language are also string type, so convert them to integer
                    Integer SOURCEID_EDU = Integer.parseInt(row_edu.getCell(EDU_Field.SOURCEID.ordinal()).getStringCellValue());
                    Integer language = Integer.parseInt(row_edu.getCell(EDU_Field.LANGUAGES.ordinal()).getStringCellValue());

                    if(SOURCEID_EN.equals(SOURCEID_EDU)) {
                        if(language.equals(2)) {
                            System.out.println("correct" + SOURCEID_EDU);
                        }else{ // some ID but language is not correct, do while
                            continue;
                        }
                    }else{
                        if(SOURCEID_EN > SOURCEID_EDU) {
                            continue;
                        }else{
                            previous_row = row_edu;
                            isNewLoop = TRUE;
                            break;
                        }
                    } 
                } // end while for education

            } // end while for entites


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

