import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

import static java.lang.Boolean.FALSE;
import static java.lang.Boolean.TRUE;
import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL;

public class EXP {
    private RESULT[] result_field = RESULT.values();

    private enum NESTED{
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

    private enum  RESULT{
        employer (3),
        department (6),
        position_title (7),
        period_start (8),
        period_end (9),
        city (4),
        country (5),
        torder (11);

        private int value;

        private RESULT(int value) {
            this.value = value;
        }

        public int getValue() {
            return this.value;
        }
    }

    public static void main(String[] args) {
        EXP exp = new EXP();
        exp.readFile("src/AH_RP.xls", "src/AH_EXP_reorganize.xlsx");
    }

    private void readFile(String FILE_PATH_1, String FILE_PATH_2) {
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
            MAIN_Field[] main_Field = MAIN_Field.values();
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
                short torder_Y = 1;
                short torder_N = 1;

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

                          // When employer, department and position_title are empty, it is invalid data.
                          // Pass this loop
                          if(!isValid(row_exp))
                              continue;

                          // If previous SOURCEID is not some to now SOURCEID, write into main sheet
                          if(!previousID.equals(SOURCEID_EN)) {
                              EDU.setMainSheet(sheet_main, row_en);
                              previousID = SOURCEID_EN;
                          }

                          String current = "Y";
                          if(row_exp.getCell(EXP_Field.current.ordinal()) != null)
                              current = row_exp.getCell(EXP_Field.current.ordinal()).getStringCellValue();

                          if(current.equals("Y")){
                              this.write_current(sheet_nested, row_exp, torder_Y);
                              torder_Y++;
                          }else{ // "N"
                              this.write_exp(sheet_nested, row_exp, torder_N);
                              torder_N++;
                          }

                          this.write_common(sheet_nested, row_en, row_exp);
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

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private boolean isValid(Row row) {
        Cell employer = row.getCell(EXP_Field.employer.ordinal(), RETURN_BLANK_AS_NULL);
        Cell department = row.getCell(EXP_Field.department.ordinal(), RETURN_BLANK_AS_NULL);
        Cell title = row.getCell(EXP_Field.position_title.ordinal(), RETURN_BLANK_AS_NULL);

        if(employer == null && department == null && title == null) {
            System.out.println("invalid data - ted_id :" + row.getCell(EXP_Field.tex_id.ordinal()).getStringCellValue());
            return FALSE;
        }
        return TRUE;
    }

    private void write_common(Sheet sheet_nested, Row row_en, Row row_exp) {
        int index = sheet_nested.getLastRowNum();
        Row row_nested = sheet_nested.getRow(index);

        row_nested.createCell(0).setCellValue(row_en.getCell(0).getStringCellValue());  // CRISID_PARENT
        row_nested.createCell(1).setCellValue(row_en.getCell(2).getStringCellValue());  // SOURCEREF_PARENT
        row_nested.createCell(2).setCellValue(row_en.getCell(3).getStringCellValue());  // SOURCEID_PARENT
        row_nested.createCell(4).setCellValue(row_en.getCell(2).getStringCellValue());  // SOURCEREF
        row_nested.createCell(5).setCellValue(row_exp.getCell(EXP_Field.tex_id.ordinal()).getStringCellValue());    // SOURCEID
    }

    private void write_exp(Sheet sheet_nested, Row row_exp, short torder_N) {
        int index = sheet_nested.getLastRowNum();
        Row row_nested = sheet_nested.createRow(index + 1);

        String pattern = "\\d{4}-\\d{2}-\\d{2}";

        for (int i = 0; i < 7; i++) {
            if(row_exp.getCell(this.result_field[i].getValue()) != null) {
                String value = row_exp.getCell(this.result_field[i].getValue()).getStringCellValue();

                // Convert type of date
                if(value.matches(pattern))
                    value = EDU.convertDate(value);

                if(value.equals("Taiwan"))
                    value = "TW";

                if (i == 6 && value.length() > 2)
                    System.out.println(row_exp.getCell(0) + "--" + value);

                if(!value.equals("\\N") && !value.equals("."))
                    row_nested.createCell(i + 14).setCellValue(value);
            }
        }

        row_nested.createCell(21).setCellValue(torder_N);
    }

    private void write_current(Sheet sheet_nested, Row row_exp, short torder_Y) {
        int index = sheet_nested.getLastRowNum();
        Row row_nested = sheet_nested.createRow(index + 1);

        String pattern = "\\d{4}-\\d{2}-\\d{2}";

        for (int i = 0; i < 7; i++) {
            if(row_exp.getCell(this.result_field[i].getValue()) != null) {
                String value = row_exp.getCell(this.result_field[i].getValue()).getStringCellValue();

                // Convert type of date
                if(value.matches(pattern))
                    value = EDU.convertDate(value);

                if(value.equals("Taiwan"))
                    value = "TW";

                if (i == 6 && value.length() > 2)
                    System.out.println(row_exp.getCell(0) + "--" + value);


                if(!value.equals("\\N") && !value.equals("."))
                    row_nested.createCell(i + 6).setCellValue(value);
            }
        }
        // torder field
        row_nested.createCell(13).setCellValue(torder_Y);
    }
}

enum EXP_Field{
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