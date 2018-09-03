import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.joda.time.LocalDate;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;
import org.joda.time.format.DateTimeFormatterBuilder;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.chrono.MinguoDate;
import java.util.*;

import static java.lang.Boolean.FALSE;
import static java.lang.Boolean.TRUE;
import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class Project {
    static String [] main_fields = {"ACTION", "CRISID", "UUID", "SOURCEREF", "SOURCEID", "logo", "title",
            "code", "pjtranslatedName", "principalinvestigator", "coinvestigators", "pjfundingorg",
            "pjorganization", "pjlink", "startdate", "expdate", "pjbugetid"};

    static String[] nested_fields = {"CRISID_PARENT", "SOURCEREF_PARENT", "SOURCEID_PARENT", "UUID",
            "SOURCEREF", "SOURCEID"};

    private void integrate_project() {
        /* ---------------------------------------start initialized--------------------------------------------------------------*/
        String path  = "src/BudgetID/AH_計畫檔_combined_add_lack_tid.xlsx";
        String sheetsName = "teacher project";
        Hashtable<String, Units> units = getUnitsTable();
        Hashtable<String, Researcher> researchers = getResearchersTable();

        try {
            InputStream PROJECT_FILE = new FileInputStream(path);
            XSSFWorkbook PROJECT_BOOK = new XSSFWorkbook(PROJECT_FILE);
            XSSFSheet project_sheet = PROJECT_BOOK.getSheet(sheetsName);

            Workbook result = new HSSFWorkbook();
            Sheet sheet_main = result.createSheet("main_entities");
            setFieldName(sheet_main, main_fields);
            Sheet sheet_nested = result.createSheet("nested_entities");
            setFieldName(sheet_nested, nested_fields);

            /* ---------------------------------------end  initialized--------------------------------------------------------------*/

            Iterator<Row> rows_project = project_sheet.iterator();
            rows_project.next();

            while (rows_project.hasNext()) {
                Row row_project = rows_project.next();

                String tid = row_project.getCell(PROJECT_ENUM.tid.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                String isuse = row_project.getCell(PROJECT_ENUM.isuse.getValue()).getStringCellValue();

                if(tid.equals("") || isuse.equals("0"))
                    continue;

                Researcher researcher = researchers.get(tid);
                if(researcher == null) {
                    // System.out.println("tid存在但是找不到 -- " + row_project.getCell(0).getStringCellValue());
                    continue;
                }

                String SOURCEID  = row_project.getCell(PROJECT_ENUM.tpjid.getValue()).getStringCellValue();
                String sub_researcher = row_project.getCell(PROJECT_ENUM.sname.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                String code = row_project.getCell(PROJECT_ENUM.pjno.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                String title = row_project.getCell(PROJECT_ENUM.title.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                String etitle = row_project.getCell(PROJECT_ENUM.etitle.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                String link  = row_project.getCell(PROJECT_ENUM.alink.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();

                String [] period = {"", ""};
                if(row_project.getCell(PROJECT_ENUM.dates.getValue()) != null){
                    // System.out.println(row_project.getCell(0).getStringCellValue());
                    period = splitPeriod(row_project.getCell(0).getStringCellValue(), row_project.getCell(PROJECT_ENUM.dates.getValue()).getStringCellValue());
                }

                String founds  = row_project.getCell(PROJECT_ENUM.founds.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                Units unit = units.get(founds);

                setResult(sheet_main, SOURCEID, title, etitle, code, researcher, sub_researcher, unit, link);
                setPeriod(sheet_main, period);

                if(row_project.getCell(PROJECT_ENUM.bugetid.getValue()) != null){
                    String start = row_project.getCell(PROJECT_ENUM.start.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                    String deadline = row_project.getCell(PROJECT_ENUM.deadline.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                    String budgetID = row_project.getCell(PROJECT_ENUM.bugetid.getValue()).getStringCellValue();

                    String org  = row_project.getCell(PROJECT_ENUM.org.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                    Units organization = units.get(org);

                    if(start.equals("") || deadline.equals("")) {
                        System.out.println("budgetID存在日期但日期不存在 -- " + row_project.getCell(0).getStringCellValue());
                    }else{
                        start = transferMinguoToAD(start, "default");
                        deadline = transferMinguoToAD(deadline, "dafult");
                        setTimeAndBudgetID(sheet_main, organization, start, deadline, budgetID);
                    }
                }

            } // end while

            FileOutputStream out = new FileOutputStream("src/BudgetID/result_project.xls");
            result.write(out);
            out.close();

        }catch (IOException e){
            e.printStackTrace();
        }
    }

    private void setPeriod(Sheet sheet_main, String[] period) {
        int index = sheet_main.getLastRowNum();
        Row row = sheet_main.getRow(index);

        int [] order = {14, 15};

        for(int i  = 0; i < order.length; i++){
            row.createCell(order[i]).setCellValue(period[i]);
        }
    }

    private String [] splitPeriod(String ID, String period) {
        period = splitComma(period);
        String [] times = {"", ""};

        if(period.matches("\\d{2,3}\\d{4}\\s*[-~]\\s*\\d{2,3}\\d{4}")) { // ex: 1060301 ~ 1070831
            times = period.split("\\s*[~-]\\s*");
            purify(times);
            times[0] = transferMinguoToAD(times[0], "default");
            times[1] = transferMinguoToAD(times[1], "default");
        }else if(period.matches("\\d{2,3}\\d{2}\\s*[-~]\\s*\\d{2,3}\\d{2}")) { // ex: 10603 ~ 10708
            times = period.split("\\s*[~-]\\s*");
            purify(times);
            times[0] = transferMinguoToAD(times[0], "min");
            times[1] = transferMinguoToAD(times[1], "max");
        }else if(period.matches("\\d*-\\d*\\s*[~]\\s*\\d*-\\d*")){ // ex 2011-08 ~ 2016-12
            times = period.split("\\s*~\\s*");

            String [] year_Month;
            String [] type = {"min", "max"};

            for(int i = 0; i < times.length; i++){
                year_Month = times[i].split("-");
                int year_int = Integer.parseInt(year_Month[0]) - 1911;
                String month = year_Month[1];
                times[i] = transferMinguoToAD(String.valueOf(year_int) + month, type[i]);
            }

        }else if(period.matches("\\d{4}/\\d{1,2}/\\d{1,2}\\s*[~-]\\s*\\d{4}/\\d{1,2}/\\d{1,2}")) { // ex 2017/1/1 ~ 2018/1/31 or 2017/01/01 ~ 2018/01/31
            times = period.split("\\s*[~-]\\s*");

            String [] year_Month_day;
            String [] type = {"min", "max"};

            for(int i = 0; i < times.length; i++){
                year_Month_day = times[i].split("/");
                int year_int = Integer.parseInt(year_Month_day[0]) - 1911;
                String month = year_Month_day[1];
                if(month.length() == 1)
                    month = "0" + month;

                times[i] = transferMinguoToAD(String.valueOf(year_int) + month, type[i]);
            }

        }else{
            System.out.println("other type :" + period);
        }

        return times;

    }

    private String splitComma(String period) {

        if(period.contains(".")){
            String [] times = period.split("\\s*[~-]\\s*");
            String [] dates = {"", ""};
            for(int i = 0; i < times.length; i++){
                    String [] year_month_day = times[i].split("\\.");
                    String result = "";

                    for(int j = 0; j < year_month_day.length; j++){
                        if(j > 0 && year_month_day[j].length() == 1){
                            result = result.concat("0" + year_month_day[j]);
                        }else {
                            result = result.concat(year_month_day[j]);
                        }
                    }
                    dates[i] = result;
            } // end for-loop
            return dates[0] + "~" + dates[1];
        }
        return period;
    }

    private void purify(String[] times) {
        for(int i = 0; i <times.length; i++){

            if(times[i].substring(0, 2).equals("20") || times[i].substring(0, 2).equals("19")){
                int year = Integer.parseInt(times[i].substring(0,4)) - 1911;
                times[i] = String.valueOf(year) + times[i].substring(4);

                if(times[i].length() == 4 && i == 0)
                {
                    times[i] = times[i] + "01";
                }else {
                    times[i] = times[i] + "12";
                }
            }
            if(times[i].substring(0,2).equals("10") && times[i].length() == 4){ // ex: 1051 --> 10501
                int month = Integer.parseInt(times[i].substring(3));
                times[i] = times[i].substring(0, 3) + "0" + month;
            }else if(times[i].substring(0,1).equals("9") && times[i].length() == 3){ // ex: 935 --> 9305
                int month = Integer.parseInt(times[i].substring(2));
                times[i] = times[i].substring(0, 2) + "0" + month;
            }

        } // end for
    }

    private static final DateTimeFormatter [] dateFormaters = {  new DateTimeFormatterBuilder()
                                                            .appendYear(4,4).appendMonthOfYear(2).appendDayOfMonth(2).toFormatter(),
                                                            new DateTimeFormatterBuilder()
                                                            .appendYear(4,4).appendMonthOfYear(2).toFormatter()};

    private String transferMinguoToAD(String date, String type) {
        DateTime time = new DateTime();

        if(date.substring(0,2).equals("10")){
            int year = Integer.parseInt(date.substring(0, 3)) + 1911;
            date = year + date.substring(3);
        }else if(date.substring(0,2).equals("11")){
            int year = Integer.parseInt(date.substring(0, 3)) + 1911;
            date = year + date.substring(3);
        }else if(date.substring(0,1).equals("9")){
            int year = Integer.parseInt(date.substring(0, 2)) + 1911;
            date = year + date.substring(2);
        }else if(date.substring(0,1).equals("8")){
            int year = Integer.parseInt(date.substring(0, 2)) + 1911;
            date = year + date.substring(2);
        }

        if(date.matches("[0-9]{8}")){   // yyyMMdd ex: 20181231
            time = dateFormaters[0].parseDateTime(date);
        }else if(date.matches("[0-9]{6}")){ // yyMMdd ex: 201510
            time = dateFormaters[1].parseDateTime(date);
        }else{
            System.out.println("TransferToAD :" + date);
        }

        if(type.equals("min")){
            time = time.withDayOfMonth(time.monthOfYear().getMinimumValue());
        }else if(type.equals("max")){
            time = time.withDayOfMonth(time.monthOfYear().getMaximumValue());
        }else {
            // default
        }

        DateTimeFormatter formatter = DateTimeFormat.forPattern("dd-MM-yyyy");
        return formatter.print(time);

    }

    private void setTimeAndBudgetID(Sheet sheet_main, Units org, String start, String deadline, String budgetID) {
        int index = sheet_main.getLastRowNum();
        Row row = sheet_main.getRow(index);

        String organization = "[CRISID=" + org.getCRISID() + "]" + org.getEn_name();
        int [] order = {12, 14, 15, 16};
        String [] fields = {organization, start, deadline, budgetID};

        for(int i  = 0; i < order.length; i++){
            row.createCell(order[i]).setCellValue(fields[i]);
        }
    }

    private void setResult(Sheet sheet_main, String SOURCEID, String title, String etitle, String code, Researcher researcher, String sub_researcher, Units unit, String link) {
       int index = sheet_main.getLastRowNum() + 1;
       Row row = sheet_main.createRow(index);

       row.createCell(0).setCellValue("CREATE");
       row.createCell(9).setCellValue("[CRISID=" + researcher.getCRISID() + "]" + researcher.getFullName());
       if(unit != null)
           row.createCell(11).setCellValue("[CRISID=" + unit.getCRISID() + "]" + unit.getEn_name());
       if(!link.equals(""))
           row.createCell(13).setCellValue("[visibility=PUBLIC URL=" + link + "]link");
       if(etitle.equals(""))
           etitle = title;

       int [] order = {4, 6, 7, 8, 10};
       String [] fields = {SOURCEID, etitle, code, title, sub_researcher};

       for(int i  = 0; i < order.length; i++){
           row.createCell(order[i]).setCellValue(fields[i]);
       }

    }

    private void setFieldName(Sheet sheet, String [] fields){
        Row row_main = sheet.createRow(0);
        for (int i = 0; i < fields.length; i++)
            row_main.createCell(i).setCellValue(fields[i]);
    }

    private void integrate_budgetID()  {
        String [] paths  = {"src/BudgetID/AH_計畫檔.xlsx", "src/BudgetID/SQL_201809.xlsx"};
        String [] sheetsName = {"teacher project", "AcademicHub"};
        try {
            InputStream PROJECT_FILE = new FileInputStream(paths[0]);
            InputStream SQL_FILE = new FileInputStream(paths[1]);

            XSSFWorkbook PROJECT_BOOK = new XSSFWorkbook(PROJECT_FILE);
            XSSFWorkbook SQL_BOOK = new XSSFWorkbook(SQL_FILE);

            XSSFSheet project_sheet = PROJECT_BOOK.getSheet(sheetsName[0]);
            XSSFSheet sql_sheet = SQL_BOOK.getSheet(sheetsName[1]);


            Iterator<Row> rows_project = project_sheet.iterator();
            rows_project.next();
            Iterator<Row> rows_sql = sql_sheet.iterator();
            rows_sql.next();

            // Control rows_dept loop
            Row previous_row = null;
            Boolean isNewLoop = FALSE;

            while (rows_project.hasNext()) {
                Row row_prject = rows_project.next();

                Cell cell_budgetID = row_prject.getCell(PROJECT_ENUM.bugetid.getValue());
                if (cell_budgetID == null) {
                    System.out.println(row_prject.getLastCellNum());
                    break;
                }
                Double budgetID_project = Double.parseDouble(cell_budgetID.getStringCellValue());

                Row row_sql = null;
                while (rows_sql.hasNext()) {
                    if(isNewLoop.equals(TRUE)) {
                        row_sql = previous_row;
                        isNewLoop = FALSE;
                    }else{
                        row_sql = rows_sql.next();
                    }

                    Double budgetID_sql = row_sql.getCell(SQL.BugetID.getValue()).getNumericCellValue();

                    if(budgetID_project.equals(budgetID_sql)) {
                        copyToPoject(row_prject, row_sql);
                    }else {
                        if(budgetID_project <= budgetID_sql) {
                            previous_row = row_sql;
                            isNewLoop = TRUE;
                            break;
                        }
                    }

                } // end while

            } // end while
            FileOutputStream out = new FileOutputStream("src/BudgetID/AH_計畫檔_combined.xlsx");
            PROJECT_BOOK.write(out);
            out.close();
        }catch (IOException e){
            e.printStackTrace();
        }
    }

    public void integrate_lack_tid(){
        String path  = "src/BudgetID/AH_RP_name.xls";
        String  sheetsName = "main_entities";
        int [] fields = {5, 3};

        Hashtable<String, String> researchers = new Hashtable<>();

        try {
            InputStream FILE = new FileInputStream(path);

            HSSFWorkbook workbook = new HSSFWorkbook(FILE);
            HSSFSheet sheet = workbook.getSheet(sheetsName);
            Iterator<Row> rows = sheet.iterator();

            Row row = rows.next();

            while(rows.hasNext()) {
                row = rows.next();

                String translatedName = row.getCell(fields[0]).getStringCellValue();
                String SOURCEID = row.getCell(fields[1]).getStringCellValue();
                researchers.put(translatedName, SOURCEID);
            } // end while

            FILE.close();

            FILE = new FileInputStream("src/BudgetID/AH_計畫檔_combined.xlsx");
            XSSFWorkbook PROJECT  = new XSSFWorkbook(FILE);
            XSSFSheet PROJECT_sheet = PROJECT.getSheet("teacher project");

            rows = PROJECT_sheet.iterator();

            row = rows.next();
            while(rows.hasNext()) {
                row = rows.next();

                String translatedName = row.getCell(PROJECT_ENUM.tname.getValue(), CREATE_NULL_AS_BLANK).getStringCellValue();
                String tid = researchers.get(translatedName);

                if(row.getCell(PROJECT_ENUM.tid.getValue()) == null) {
                    row.createCell(PROJECT_ENUM.tid.getValue()).setCellValue(tid);
                }
            } // end while

            FileOutputStream out = new FileOutputStream("src/BudgetID/AH_計畫檔_combined_add_lack_tid.xlsx");
            PROJECT.write(out);
            out.close();
        }catch (IOException e){
            e.printStackTrace();
        }
    }

    private void copyToPoject(Row row_prject, Row row_sql) {
        int index = row_prject.getLastCellNum();

        for(int i = 0; i < 8; i++){
            Cell cell = row_prject.createCell(index++);
            Cell cell_sql = row_sql.getCell(i + 1);
            String value = "";
            if(cell_sql != null)
                value = cell_sql.getStringCellValue();
            cell.setCellValue(value);
        }

    }

    private static Hashtable<String, Units> getUnitsTable() {
        String [] paths  = {"src/BudgetID/AH_dep_default.xlsx", "src/BudgetID/result_org_edit_add.xls"};
        String [] sheetsName = {"dep_code", "main_entities"};
        List<int []> fields = new ArrayList<>();
        int [] dep_default = {0, 4, 5};
        int [] org_edit = {0, 4, 5};
        fields.add(dep_default);
        fields.add(org_edit);

        Hashtable<String, Units> units = new Hashtable<>();

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

                    String CRISID = row.getCell(oreder[0]).getStringCellValue();
                    String en_name = row.getCell(oreder[1]).getStringCellValue();
                    String ch_name = row.getCell(oreder[2]).getStringCellValue();

                    units.put(ch_name, new Units(CRISID, en_name));
                } // end while

            } // end for
            units.put("國立臺灣大學", new Units( "ou00001", "National Taiwan University"));

            return units;

        }catch (IOException e){
            e.printStackTrace();
        }
        return units;
    }

    private static Hashtable<String, Researcher> getResearchersTable() {
        String [] paths  = {"src/BudgetID/AH_RP_name.xls"};
        String [] sheetsName = {"main_entities"};
        List<int []> fields = new ArrayList<>();
        int [] field_order = {0, 4, 5};
        fields.add(field_order);

        Hashtable<String, Researcher> researchers = new Hashtable<>();

        try {
            for (int i  = 0 ; i < paths.length; i++){
                InputStream FILE = new FileInputStream(paths[i]);
                Iterator<Row> rows;

                HSSFWorkbook workbook = new HSSFWorkbook(FILE);
                HSSFSheet sheet = workbook.getSheet(sheetsName[i]);
                rows = sheet.iterator();

                Row row = rows.next();

                int [] oreder = fields.get(i);
                while(rows.hasNext()) {
                    row = rows.next();

                    String SOURCEID = row.getCell(3).getStringCellValue();
                    String CRISID = row.getCell(oreder[0]).getStringCellValue();
                    String fullName = row.getCell(oreder[1]).getStringCellValue();
                    String translatedName = row.getCell(oreder[2]).getStringCellValue();

                    researchers.put(SOURCEID, new Researcher(CRISID, fullName, translatedName));
                } // end while

            } // end for

            return researchers;

        }catch (IOException e){
            e.printStackTrace();
        }
        return researchers;
    }


    public static void main(String[] args) {
        Project project = new Project();
        project.integrate_project();
        // System.out.println("10308".substring(4,5));
        // System.out.println("10308".concat("31"));
        // System.out.println("105.07.31".split("\\.").length);
    }

}

enum PROJECT_ENUM{
    tpjid(0),
    tid(1),
    pjno(2),
    title(3),
    etitle(4),
    tname(5),
    sname(6),
    dates(7),
    alink(8),
    founds(9),
    upddate(10),
    isuse(11),
    bugetid(12),
    unintc(14),
    org(17),
    start(19),
    deadline(20);


    private int value;

    private PROJECT_ENUM(int value) {
        this.value = value;
    }

    public int getValue() {
        return this.value;
    }
}

enum SQL{
    BugetID(0),
    ProjectName(1),
    unitc(2),
    Leader(3),
    seq(4),
    BugetDepName(5),
    ProjectNo(6),
    Start(7),
    Deadline(8);

    private int value;

    private SQL(int value) {
        this.value = value;
    }

    public int getValue() {
        return this.value;
    }
}