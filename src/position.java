import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Hashtable;
import java.util.Iterator;

import static java.lang.Boolean.FALSE;
import static java.lang.Boolean.TRUE;

public class position {
    private String [] main_fields = {"ACTION", "CRISID", "UUID", "SOURCEREF", "SOURCEID", "dept", "rptitle"};
    private String[] nested_fields = {"CRISID_PARENT", "SOURCEREF_PARENT", "SOURCEID_PARENT", "UUID", "SOURCEREF", "SOURCEID", "mainunit", "mainunit_role"};

    private enum DEPT_CODE{
        CRISID (0),
        UUID (1),
        SOURCEREF (2),
        SOURCEID (3),
        name (4);

        private int value;

        private DEPT_CODE(int value) {
            this.value = value;
        }

        public int getValue() {
            return this.value;
        }
    }

    private enum DEPT_Field {
        tdid (0),
        tid (1),
        dept (2),
        title (3),
        seq (4);

        private int value;

        private DEPT_Field(int value) {
            this.value = value;
        }

        public int getValue() {
            return this.value;
        }
    }

    private enum TITLE_Field {
        pcode (0),
        ename (1);

        private int value;

        private TITLE_Field(int value) {
            this.value = value;
        }

        public int getValue() {
            return this.value;
        }
    }

    private void readFile(String FILE_PATH_1, String FILE_PATH_2) {
        try {
            InputStream FILE_1 = new FileInputStream(FILE_PATH_1);
            InputStream FILE_2 = new FileInputStream(FILE_PATH_2);

            HSSFWorkbook RP = new HSSFWorkbook(FILE_1);
            XSSFWorkbook department = new XSSFWorkbook(FILE_2);

            Hashtable<String, Dept> dept_table = this.getDeptTable(department.getSheet("dep_code"));
            Hashtable<String, String> title_table = this.getTitleTable(department.getSheet("Title_code"));

            Workbook result = new HSSFWorkbook();
            Sheet sheet_main = result.createSheet("main_entities");
            Sheet sheet_nested = result.createSheet("nested_entities");

            HSSFSheet entites = RP.getSheet("main_entities");
            XSSFSheet RP_dept = department.getSheet("RP_dept");

            Iterator<Row> rows_en = entites.iterator();
            Iterator<Row> rows_dept = RP_dept.iterator();

            // Set first row in main sheet
            rows_en.next();
            setFirstRow(sheet_main, main_fields);

            // Set first row in nested sheet
            rows_dept.next();
            setFirstRow(sheet_nested, nested_fields);

            // Control rows_dept loop
            Row previous_row = null;
            Boolean isNewLoop = FALSE;
            Integer previousID = 0;

            while(rows_en.hasNext()){
                Row row_en = rows_en.next();
                // Get source ID from AH_RP.xls and compare to education sheet
                Integer SOURCEID_EN = Integer.parseInt(row_en.getCell(RP_Field.SOURCEID.ordinal()).getStringCellValue());


                Row row_dept = null;

                while (rows_dept.hasNext()){
                    if(isNewLoop.equals(TRUE)) {
                        row_dept = previous_row;
                        isNewLoop = FALSE;
                    }else{
                        row_dept = rows_dept.next();
                    }

                    // Get source ID and language
                    Integer SOURCEID_DEPT = Integer.parseInt(row_dept.getCell(DEPT_Field.tid.getValue()).getStringCellValue());

                    String dept_name = getDeptName(row_dept, dept_table);
                    String title = getTitleName(row_dept, title_table);

                    // Error row
                    if (dept_name.equals("") || title.equals("")){
                        if(!previousID.equals(SOURCEID_EN))
                            System.out.println("SOURCEID:" + SOURCEID_DEPT);
                    }

                    if(SOURCEID_EN.equals(SOURCEID_DEPT)) {

                        if(!dept_name.equals("") && !title.equals("")){ // Correct Row
                            // If previous SOURCEID is not some to now SOURCEID, write into main sheet
                            if (!previousID.equals(SOURCEID_EN)) {
                                EDU.setMainSheet(sheet_main, row_en);
                                set_rptitle(sheet_main, title);
                                previousID = SOURCEID_EN;
                            }

                            Dept dept = getDept(row_dept, dept_table);
                            setNested(sheet_nested, row_en, row_dept, dept, title);
                            updateDept(sheet_main, dept);
                        }

                    }else {
                        if(SOURCEID_EN <= SOURCEID_DEPT) {
                            previous_row = row_dept;
                            isNewLoop = TRUE;
                            break;
                        }
                    }

                } // end while for departments

            } // end while for entities

            FileOutputStream out = new FileOutputStream("result_dep.xls");
            result.write(out);
            out.close();

        } catch (IOException e){
            e.printStackTrace();
        }
    }

    private Dept getDept(Row row_dept, Hashtable<String,Dept> dept_table) {
        String dept_code = row_dept.getCell(DEPT_Field.dept.getValue()).getStringCellValue();
        Dept dept = dept_table.get(dept_code);

        return  dept;
    }

    private String getTitleName(Row row_dept, Hashtable<String, String> title_table) {
        String title_code = row_dept.getCell(DEPT_Field.title.getValue()).getStringCellValue();

        String title = "";
        if (title_table.get(title_code) != null)
            title = title_table.get(title_code);

        return title;
    }

    private String getDeptName(Row row_dept, Hashtable<String, Dept> dept_table) {
        String dept_code = row_dept.getCell(DEPT_Field.dept.getValue()).getStringCellValue();
        Dept dept = dept_table.get(dept_code);

        String dept_name = "";
        // Check valid row
        if (dept != null)
            dept_name = dept.getName();

        return dept_name;
    }

    private void set_rptitle(Sheet sheet, String title) {
        int index = sheet.getLastRowNum();
        Row row = sheet.getRow(index);
        int orderOfTitle = 6;

        row.createCell(orderOfTitle).setCellValue(title);
    }

    private void updateDept(Sheet sheet, Dept dept) {
        int index = sheet.getLastRowNum();
        Row row = sheet.getRow(index);
        int orderOfDept = 5;

        if (row.getCell(orderOfDept) == null){
            Cell cell = row.createCell(orderOfDept);
            cell.setCellValue("[CRISID=" + dept.getCRISID() + "]" + dept.getName());
        }else{
            Cell cell = row.getCell(orderOfDept);
            String orig_dept = cell.getStringCellValue();
            String new_dept = "[CRISID=" + dept.getCRISID() + "]" + dept.getName();

            if(!orig_dept.contains(new_dept))
                cell.setCellValue(new_dept + "|||" + orig_dept);
        }
    }

    private void setNested(Sheet sheet, Row row_en, Row row_dept, Dept dept, String title) {
        int index = sheet.getLastRowNum();
        Row row = sheet.createRow(index + 1);

        String CRISID_PARENT = row_en.getCell(RP_Field.CRISID.ordinal()).getStringCellValue();
        String SOURCEID_PARENT = row_en.getCell(RP_Field.SOURCEID.ordinal()).getStringCellValue();
        String SOURCEID = row_dept.getCell(DEPT_Field.tdid.getValue()).getStringCellValue();
        String mainunit = "[CRISID=" + dept.getCRISID() + "]" + dept.getName();

        String [] nested = {CRISID_PARENT, "AH", SOURCEID_PARENT, "", "AH", SOURCEID, mainunit, title};

        for (int i = 0; i < nested.length; i++)
            row.createCell(i).setCellValue(nested[i]);
    }

    public static void setFirstRow(Sheet sheet, String[] fields) {
        Row row_main = sheet.createRow(0);
        for (int i = 0; i < fields.length; i++)
            row_main.createCell(i).setCellValue(fields[i]);
    }

    private  Hashtable<String, String>getTitleTable(XSSFSheet sheet) {
        // Create table
        Hashtable<String, String> title_table = new Hashtable<String, String>();

        // Skip first row in dept_code sheet
        Iterator<Row> rows = sheet.iterator();
        rows.next();

        while (rows.hasNext()) {
            Row row = rows.next();

            String pcode = row.getCell(TITLE_Field.pcode.getValue()).getStringCellValue();
            String ename = row.getCell(TITLE_Field.ename.getValue()).getStringCellValue();

            title_table.put(pcode, ename);
        }

        return title_table;
    }

    private Hashtable<String, Dept> getDeptTable(XSSFSheet sheet) {
        // Create table
        Hashtable<String, Dept> dept_table = new Hashtable<String, Dept>();

        // Skip first row in dept_code sheet
        Iterator<Row> rows = sheet.iterator();
        rows.next();

        while (rows.hasNext()) {
            Row row = rows.next();

            String CRISID = row.getCell(DEPT_CODE.CRISID.getValue()).getStringCellValue();
            String UUID = row.getCell(DEPT_CODE.UUID.getValue()).getStringCellValue();
            String SOURCEID = row.getCell(DEPT_CODE.SOURCEID.getValue()).getStringCellValue();
            String name = row.getCell(DEPT_CODE.name.getValue()).getStringCellValue();

            Dept dept = new Dept(CRISID, UUID, name);
            dept_table.put(SOURCEID, dept);
        }

        return  dept_table;
    }

    public static void main(String[] args) {
        position pos = new position();
        pos.readFile("src/AH_RP.xls", "src/AH_dep.xlsx");
    }
}
