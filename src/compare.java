import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
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

    public enum NESTED{
        CRISID_PARENT,
        SOURCEREF_PARENT,
        SOURCEID_PARENT,
        UUID,
        OURCEREF,
        SOURCEID,
        eduno,
        eduschool,
        edumajor,
        edudegree,
        edustart,
        eduend,
        educity,
        educountry
    }

    public static void showMemoryUsage()
    {
        long memory = Runtime.getRuntime().totalMemory()
                - Runtime.getRuntime().freeMemory();
        System.out.println(String.format("%.1f MB", (memory / (1024.0 * 1024.0))));
    }

    public static void readFile(String FILE_PATH_1, String FILE_PATH_2){
        try {
            InputStream FILE_1 = new FileInputStream(FILE_PATH_1);
            InputStream FILE_2 = new FileInputStream(FILE_PATH_2);

            HSSFWorkbook RP = new HSSFWorkbook(FILE_1);
            XSSFWorkbook EDU = new XSSFWorkbook(FILE_2);

            Workbook result = new HSSFWorkbook();
            Sheet sheet_main = result.createSheet("main_entities");
            Sheet sheet_nested = result.createSheet("nested_entities");

            HSSFSheet entites = RP.getSheet("main_entities");
            XSSFSheet educations = EDU.getSheet("education");

            Iterator<Row> rows_en = entites.iterator();
            Iterator<Row> rows_edu = educations.iterator();

            // Set first row in main sheet
            rows_en.next();
            Row row_main = sheet_main.createRow(0);
            RP_Field[] main_Field = RP_Field.values();
            for (int i = 0; i < main_Field.length; i++) {
                row_main.createCell(i).setCellValue(main_Field[i].toString());
            }

            // Set first row in nested sheet
            rows_edu.next();
            Row row_nested = sheet_nested.createRow(0);
            NESTED[] nested_Field = NESTED.values();
            for (int i = 0; i < nested_Field.length; i++) {
                row_nested.createCell(i).setCellValue(nested_Field[i].toString());
            }

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
                        if(language.equals(2)) { // Correct row

                            int index_main = sheet_main.getLastRowNum(); // Get current number of Rows
                            // Get Previous SOURCEID
                            Integer previousID = 0;
                            if(previous_row != null) {
                                String _previousID = row_main.getCell(RP_Field.SOURCEID.ordinal()).getStringCellValue();
                                previousID = Integer.parseInt(_previousID);
                            }
                            // If previous SOURCEID is not some to now SOURCEID, write into main sheet
                            if(!previousID.equals(SOURCEID_EN)) {
                                row_main = sheet_main.createRow(index_main + 1);
                                for (int i = 0; i < 4; i++) {
                                    row_main.createCell(i).setCellValue(row_en.getCell(i).getStringCellValue());
                                }
                            }

                            // Write into nested sheet
                            int index_nested = sheet_nested.getLastRowNum(); // Get current of Rows
                            row_nested = sheet_nested.createRow(index_nested + 1);

                            row_nested.createCell(0).setCellValue(row_en.getCell(0).getStringCellValue());
                            row_nested.createCell(1).setCellValue(row_en.getCell(2).getStringCellValue());
                            row_nested.createCell(2).setCellValue(row_en.getCell(3).getStringCellValue());
                            row_nested.createCell(4).setCellValue(row_en.getCell(2).getStringCellValue());
                            row_nested.createCell(5).setCellValue(row_edu.getCell(0).getStringCellValue());
                            row_nested.createCell(6).setCellValue("1");
                            System.out.print(SOURCEID_EDU);
                            if(row_edu.getCell(3) == null){
                                System.out.print("-- Major");
                            }
                            if(row_edu.getCell(6) == null){
                                System.out.print("-- Major");
                            }
                            if(row_edu.getCell(7) == null){
                                System.out.print("-- Degree");
                            }
                            if(row_edu.getCell(8) == null){
                                System.out.print("-- Start");
                            }
                            if(row_edu.getCell(9) == null){
                                System.out.print("-- End");
                            }
                            if(row_edu.getCell(4) == null){
                                System.out.print("-- City");
                            }
                            System.out.println();
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

            FileOutputStream out = new FileOutputStream("result.xls");
            result.write(out);
            out.close();

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

