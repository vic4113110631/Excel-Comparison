import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

import static java.lang.Boolean.FALSE;
import static java.lang.Boolean.TRUE;

public class EXP {
    public enum RP_Field{
        CRISID,
        UUID,
        SOURCEREF,
        SOURCEID
    }

    public enum EXP_Field{
        tex_id,
        tid,
        languages,
        employer,
        city,
        country,
        department,
        position_title,
        period_start,
        period_end,
        current,
        torder
    }

    public enum NESTED{
        CRISID_PARENT,
        SOURCEREF_PARENT,
        SOURCEID_PARENT,
        UUID,
        SOURCEREF,
        SOURCEID,
        rpcur_employer,
        rpcur_depart,
        rpcur_title,
        rpcur_start,
        rpcur_end,
        rpcur_city,
        rpcur_country,
        rpcur_order,
        rpexp_employer,
        rpexp_depart,
        rpexp_title,
        rpexp_start,
        rpexp_end,
        rpexp_city,
        rpexp_country,
        rpexp_order,
    }

    public static void main(String[] args) {
        readFile("src/AH_RP.xls", "src/AH_EXP.xlsx");
    }

    private static void readFile(String FILE_PATH_1, String FILE_PATH_2) {
        try {
            InputStream FILE_1 = new FileInputStream(FILE_PATH_1);
            InputStream FILE_2 = new FileInputStream(FILE_PATH_2);

            HSSFWorkbook RP = new HSSFWorkbook(FILE_1);
            XSSFWorkbook EXP = new XSSFWorkbook(FILE_2);

            Workbook result = new HSSFWorkbook();
            Sheet sheet_main = result.createSheet("main_entities");
            Sheet sheet_nested = result.createSheet("nested_entities");

            HSSFSheet entites = RP.getSheet("main_entities");
            XSSFSheet exp = EXP.getSheet("teacher_experience");

            Iterator<Row> rows_en = entites.iterator();
            Iterator<Row> rows_exp = exp.iterator();

            // Set first row in main sheet
            rows_en.next();
            Row row_main = sheet_main.createRow(0);
            RP_Field[] main_Field = RP_Field.values();
            for (int i = 0; i < main_Field.length; i++) {
                row_main.createCell(i).setCellValue(main_Field[i].toString());
            }

            // Set first row in nested sheet
            rows_exp.next();
            Row row_nested = sheet_nested.createRow(0);
            NESTED[] nested_Field = NESTED.values();
            for (int i = 0; i < nested_Field.length; i++) {
                row_nested.createCell(i).setCellValue(nested_Field[i].toString());
            }

            // Control rows_exp loop
            Row previous_row = null;
            Boolean isNewLoop = FALSE;
            Integer previousID = 0;

            while (rows_en.hasNext()) {
                Row row_en = rows_en.next();
                // Get source ID from AH_RP.xls and compare to education sheet
                Integer SOURCEID_EN = Integer.parseInt(row_en.getCell(RP_Field.SOURCEID.ordinal()).getStringCellValue());

                Row row_exp = null;
                while (rows_exp.hasNext()){
                    if(isNewLoop.equals(TRUE)) {
                        row_exp = previous_row;
                        isNewLoop = FALSE;
                    }else{
                        row_exp = rows_exp.next();
                    }

                    // Get source ID and language
                    Integer SOURCEID_EXP = Integer.parseInt(row_exp.getCell(EXP_Field.tid.ordinal()).getStringCellValue());
                    Integer language = Integer.parseInt(row_exp.getCell(EXP_Field.languages.ordinal()).getStringCellValue());

                    if(SOURCEID_EN.equals(SOURCEID_EXP)) {
                      if(language.equals(2)){
                          System.out.println("correct:" + SOURCEID_EXP);

                          // If previous SOURCEID is not some to now SOURCEID, write into main sheet
                          int index_main = sheet_main.getLastRowNum(); // Get current number of Rows
                          if(!previousID.equals(SOURCEID_EN)) {
                              row_main = sheet_main.createRow(index_main + 1);
                              for (int i = 0; i < 4; i++) {
                                  row_main.createCell(i).setCellValue(row_en.getCell(i).getStringCellValue());
                              }
                              previousID = Integer.parseInt(row_main.getCell(RP_Field.SOURCEID.ordinal()).getStringCellValue());
                          }

                      }
                    }else{
                        if(SOURCEID_EN <= SOURCEID_EXP) {
                            previous_row = row_exp;
                            isNewLoop = TRUE;
                            break;
                        }
                    }
                } // end while for experience

            } // end while for entities

            FileOutputStream out = new FileOutputStream("result_exp.xls");
            result.write(out);
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
