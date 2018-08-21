import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL;

public class Unit {
    private String [] main_fields = {"ACTION", "CRISID", "UUID", "SOURCEREF", "SOURCEID",
                                     "name", "orgTranslatedName", "description", "city", "iso-3166-country"};

    private List<String> exclude = new ArrayList<>(Arrays.asList("test", "."));

    public static void main(String[] args) {
        Unit unit = new Unit();
        String [] paths  = {"src/AH_計畫檔.xlsx", "src/Project_SQL_201809.xlsx"};
        String [] sheetsName = {"teacher project", "AcademicHub"};
        String [] fields = {"founds", "BugetDepName"};
        unit.unify(paths, sheetsName, fields);
    }

    private void unify(String [] paths, String [] sheetsName, String [] fields) {
        try {
            Workbook result = new HSSFWorkbook();
            Sheet sheet_result = result.createSheet("main_entities");
            position.setFirstRow(sheet_result, main_fields);

            List<String> units = new ArrayList<String>();

            for (int i  = 0 ; i < paths.length; i++){
                InputStream FILE = new FileInputStream(paths[i]);
                XSSFWorkbook workbook = new XSSFWorkbook(FILE);
                XSSFSheet sheet = workbook.getSheet(sheetsName[i]);

                Iterator<Row> rows = sheet.iterator();

                // Get order of field
                Row row = rows.next();
                int order = 0;
                for (Cell cell : row) {
                    if(cell.getStringCellValue().trim().matches(fields[i]))
                        break;
                    order++;
                }

                while(rows.hasNext()) {
                    row = rows.next();

                    String unit_name;
                    Cell cell = row.getCell(order, RETURN_BLANK_AS_NULL);

                    if (cell == null){
                        // printError(row, paths[i]);
                        continue;
                    }else{
                        unit_name = row.getCell(order).getStringCellValue().trim();

                        if (exclude.contains(unit_name)) {
                            printError(row, paths[i]);
                            continue;
                        }
                        // Unit Name likes NSC99-2628-B002-013-MY3 is error.
                        if(count(unit_name, '-') > 3) {
                            printError(row, paths[i]);
                            continue;
                        }
                    }
                    // If unit name Non-repeat, set the result.
                    if(!units.contains(unit_name)){
                        units.add(unit_name);
                        setResult(sheet_result, unit_name);
                    }

                } // end while

            } // end for

            FileOutputStream out = new FileOutputStream("result_org.xls");
            result.write(out);
            out.close();

        }catch (IOException e){
            e.printStackTrace();
        }
    }

    private void printError(Row row, String path) {
        CellType type = row.getCell(0).getCellTypeEnum();

        System.out.print(path.replace("src/","") + "--");
        if(type.equals(CellType.STRING)) {
            System.out.println("ID:" + row.getCell(0).getStringCellValue());
        }else{
            double value = row.getCell(0).getNumericCellValue();
            System.out.printf("ID:%.0f", value);
            System.out.println();
        }
    }

    private void setResult(Sheet sheet, String unit_name) {
        int index = sheet.getLastRowNum();
        Row row = sheet.createRow(index + 1);

        row.createCell(6).setCellValue(unit_name);
    }

    public static int count(String word, Character ch)
    {
        int pos = word.indexOf(ch);
        return pos == -1 ? 0 : 1 + count(word.substring(pos+1),ch);
    }
}
