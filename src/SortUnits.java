
import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL;

public class SortUnits {
    private String [] main_fields = {"ACTION", "CRISID", "UUID", "SOURCEREF", "SOURCEID",
                                     "name", "orgTranslatedName", "description", "city", "iso-3166-country"};
    private List<String> synonym = new ArrayList<>(Arrays.asList("科技部", "MOST", "Ministry of Science and Technology",
                                                               "行政院國家科學委員會", "國科會", "National Science Council"));
    private List<String> exclude = new ArrayList<>(Arrays.asList("test", "."));

    public static void main(String[] args) {
        SortUnits unit = new SortUnits();
        //String [] paths  = {"src/AH_計畫檔.xlsx", "src/Project_SQL_201809.xlsx"};
        //String [] sheetsName = {"teacher project", "AcademicHub"};
        //String [] fields = {"founds", "BugetDepName"};

        //unit.unify(paths, sheetsName, fields);

        String [] paths  = {"src/AH_dep_default.xlsx", "src/BudgetID/result_org_edit_add.xls"};
        String [] sheetsName = {"dep_code", "main_entities"};
        List<int []> fields = new ArrayList<>();
        int [] dep_default = {4, 5};
        int [] org_edit = {5, 6};
        fields.add(dep_default);
        fields.add(org_edit);

        Hashtable<String, String> units = SortUnits.check(paths, sheetsName, fields);
        units.put("國立臺灣大學", "National Taiwan University");
        SortUnits.findNotInKey(units);
        /*
        for (String key : units.keySet()) {
            String value = units.get(key);
            System.out.println(key + "--" + value);
        }

        System.out.println(units.size());
        */
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
                        System.out.print("欄位為空--");
                        printError(row, paths[i]);
                        continue;
                    }else{
                        unit_name = row.getCell(order).getStringCellValue().trim();

                        if (exclude.contains(unit_name)) {
                            System.out.print("欄位為"+ unit_name + "--");
                            printError(row, paths[i]);
                            continue;
                        }
                        // Unit Name likes NSC99-2628-B002-013-MY3 is error.
                        if(count(unit_name, '-') > 1) {
                            System.out.print("欄位為計畫編號--");
                            printError(row, paths[i]);
                            continue;
                        }

                        if(synonym.contains(unit_name)){
                            if (!unit_name.equals("科技部") && !unit_name.equals("國科會") && !unit_name.equals("行政院國家科學委員會")) {
                                System.out.print("欄位為科技部或國科會英文--" + unit_name + "--");
                                printError(row, paths[i]);
                            }
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

    private void removeRedundant(String path, String sheet_name){
        try {
            InputStream FILE = new FileInputStream(path);
            XSSFWorkbook workbook = new XSSFWorkbook(FILE);
            Sheet sheet = workbook.getSheet(sheet_name);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if(cell.getCellTypeEnum().equals(CellType.STRING)){
                        String value = cell.getStringCellValue().trim();
                        value = StringEscapeUtils.unescapeHtml4(value);
                        if(value.equals("."))
                            value = "";
                        cell.setCellValue(value);
                    }
                } // end for loop

            } // end for loop


            FILE.close();

            FileOutputStream output  = new FileOutputStream(new File("src/AH_計畫檔_reorganize.xlsx"));
            workbook.write(output);
            output.close();


        }catch (IOException e){
            e.printStackTrace();
        }
    }

    public static Hashtable<String, String> check(String [] paths, String [] sheetsName, List<int []> fields) {
        Hashtable<String, String> units = new Hashtable<>();

        try {
            for (int i  = 0 ; i < paths.length; i++){
                InputStream FILE = new FileInputStream(paths[i]);
                Iterator<Row> rows;
                if (paths[i].contains("xlsx")) {
                    XSSFWorkbook workbook = new XSSFWorkbook(FILE);
                    XSSFSheet sheet = workbook.getSheet(sheetsName[i]);
                    rows = sheet.iterator();
                }else{
                    HSSFWorkbook workbook = new HSSFWorkbook(FILE);
                    HSSFSheet sheet = workbook.getSheet(sheetsName[i]);
                    rows = sheet.iterator();
                }

                Row row = rows.next();

                int [] oreder = fields.get(i);
                while(rows.hasNext()) {
                    row = rows.next();

                    String en_name = row.getCell(oreder[0]).getStringCellValue();
                    String ch_name = row.getCell(oreder[1]).getStringCellValue();

                    units.put(ch_name, en_name);
                } // end while

            } // end for

            return units;

        }catch (IOException e){
            e.printStackTrace();
        }
        return units;
    } // end method

    public static void findNotInKey(Hashtable<String, String> units) {
        String [] paths  = {"src/BudgetID/AH_計畫檔.xlsx"};
        String [] sheetsName = {"teacher project"};
        String [] fields = {"founds"};

        try {
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
                        continue;
                    }else{
                        String ID = row.getCell(0).getStringCellValue();

                        if(cell.getCellTypeEnum().equals(CellType.STRING)) {
                            unit_name = row.getCell(order).getStringCellValue().trim();

                            if(!units.containsKey(unit_name)) {
                                System.out.println(unit_name + "--" + ID);
                            }
                        }else{
                            Double value = row.getCell(order).getNumericCellValue();
                            System.out.println(value + ID);
                        }

                    }
                } // end while

            } // end for

        }catch (IOException e){
            e.printStackTrace();
        }

    }
}
