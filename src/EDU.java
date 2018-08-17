import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import static java.lang.Boolean.FALSE;
import static java.lang.Boolean.TRUE;
import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL;


public class EDU {

    private enum  EDU_Field{
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

    enum NESTED{
        CRISID_PARENT,
        SOURCEREF_PARENT,
        SOURCEID_PARENT,
        UUID,
        SOURCEREF,
        SOURCEID,
        eduno,
        eduschool,
        edumajor,
        edudegree,
        edustart,
        eduend,
        educity,
        educountry;
    }

    private static void readFile(String FILE_PATH_1, String FILE_PATH_2){
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
            MAIN_Field[] main_Field = MAIN_Field.values();
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
            Integer previousID = 0;

            while (rows_en.hasNext()) {
                Row row_en = rows_en.next();
                // get source ID from AH_RP.xls and compare to education sheet
                // default source ID from AH_RP.xls is string, so convert it to integer
                Integer SOURCEID_EN = Integer.parseInt(row_en.getCell(RP_Field.SOURCEID.ordinal()).getStringCellValue());

                Row row_edu = null;
                short torder = 1;

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

                            // When school_name, major and degree are empty, it is invalid data.
                            // Pass this loop
                            if(!isValid(row_edu))
                                continue;

                            // If previous SOURCEID is not some to now SOURCEID, write into main sheet
                            if(!previousID.equals(SOURCEID_EN)) {
                                setMainSheet(sheet_main, row_en);
                                previousID = SOURCEID_EN;   // Record previous ID to avoid same data write in main sheet
                            }else{
                                torder++;
                            }

                            // Write into nested sheet
                            int index_nested = sheet_nested.getLastRowNum(); // Get current of Rows
                            row_nested = sheet_nested.createRow(index_nested + 1);

                            row_nested.createCell(0).setCellValue(row_en.getCell(0).getStringCellValue());  // CRISID_PARENT
                            row_nested.createCell(1).setCellValue(row_en.getCell(2).getStringCellValue());  // SOURCEREF_PARENT
                            row_nested.createCell(2).setCellValue(row_en.getCell(3).getStringCellValue());  // SOURCEID
                            row_nested.createCell(4).setCellValue(row_en.getCell(2).getStringCellValue());  // SOURCERECH
                            row_nested.createCell(5).setCellValue(row_edu.getCell(0).getStringCellValue()); // SOURCE
                            row_nested.createCell(6).setCellValue(torder);

                            setNested(row_nested, row_edu);

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

            } // end while for entities

            FileOutputStream out = new FileOutputStream("result.xls");
            result.write(out);
            out.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean isValid(Row row_edu) {
        Cell degree = row_edu.getCell(EDU_Field.DEGREE.ordinal(), RETURN_BLANK_AS_NULL);
        Cell school = row_edu.getCell(EDU_Field.SCHOOL_NAME.ordinal(), RETURN_BLANK_AS_NULL);
        Cell major = row_edu.getCell(EDU_Field.MAJOR.ordinal(), RETURN_BLANK_AS_NULL);

        if(degree == null && school == null && major == null) {
            System.out.println("invalid data - ted_id :" + row_edu.getCell(EDU_Field.TED_ID.ordinal()).getStringCellValue());
            return FALSE;
        }
        return TRUE;
    }

    static void setMainSheet(Sheet sheet_main, Row row_en) {
        int index_main = sheet_main.getLastRowNum(); // Get current number of Rows
        Row row = sheet_main.createRow(index_main + 1);

        row.createCell(0).setCellValue("UPDATE");
        for (int i = 0; i < 4; i++) {
            row.createCell(i + 1).setCellValue(row_en.getCell(i).getStringCellValue());
        }
    }

    private static void setNested(Row row_nested, Row row_edu) {
        if(row_edu.getCell(3) != null){ // School Name
            row_nested.createCell(7).setCellValue(row_edu.getCell(3).getStringCellValue());
        }
        if(row_edu.getCell(6) != null){ // Major
            row_nested.createCell(8).setCellValue(row_edu.getCell(6).getStringCellValue());
        }
        if(row_edu.getCell(7) != null){ // Degree
            row_nested.createCell(9).setCellValue(row_edu.getCell(7).getStringCellValue());
        }

        if(row_edu.getCell(8) != null){ // Start
            String date = convertDate(row_edu.getCell(8).getStringCellValue());
            // If date is default, write empty value in cell.
            if (date.equals("31-12-1899"))
                date = "";
            row_nested.createCell(10).setCellValue(date);
        }
        if(row_edu.getCell(9) != null){  // End
            String date = convertDate(row_edu.getCell(9).getStringCellValue());
            row_nested.createCell(11).setCellValue(date);
        }

        if(row_edu.getCell(4) != null){  // City
            row_nested.createCell(12).setCellValue(row_edu.getCell(4).getStringCellValue());
        }
        if(row_edu.getCell(5) != null){  // Country
            row_nested.createCell(13).setCellValue(row_edu.getCell(5).getStringCellValue());
        }
    }

    static String convertDate(String date) {
        SimpleDateFormat inputFormat = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat outputFormat = new SimpleDateFormat("dd-MM-yyyy");

        try {
            return outputFormat.format(inputFormat.parse(date));
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return date;
    }

    public static void main(String[] args) {
        readFile("src/AH_RP.xls", "src/AH_Education.xlsx");
    }
}

enum RP_Field{
    CRISID,
    UUID,
    SOURCEREF,
    SOURCEID
}

enum MAIN_Field{
    ACTION,
    CRISID,
    UUID,
    SOURCEREF,
    SOURCEID
}
