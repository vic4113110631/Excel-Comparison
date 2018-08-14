import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Reorganize {
    private Sheet reorganize(Sheet sheet){
        Iterator<Row> rows = sheet.iterator();
        // This List used to save the row that wait to be removed
        List<Integer> remove = new ArrayList<>();


        while (rows.hasNext()){
            Row row = rows.next();

            if (row.getCell(EXP_Field.current.ordinal()) == null) { // complete break lines
                Row next = rows.next();

                removeBackSlash(row);
                concatBrokenRow(row, next);

                remove.add(next.getRowNum());
                System.out.println( "--" + next.getRowNum());
            }

        } // end while

        // Delete redundant rows
        int shift = 0; // row number will change after deleting, so using a variable to control
        for (Integer i: remove) {
            sheet.shiftRows(i + 1 + shift ,sheet.getLastRowNum(), -1);
            shift--;
        }

        return sheet;
    }

    private void concatBrokenRow(Row row, Row next) {
        int numberOfColumns = row.getLastCellNum();

        // Get broken cells between two rows
        Cell last = row.getCell(numberOfColumns - 1);
        Cell broken = next.getCell(0);

        // contact two string
        last.setCellValue(last.getStringCellValue().concat(" " + broken.getStringCellValue()));

        Iterator<Cell> cells = next.cellIterator();
        cells.next();   // Skip first column that just be contacted
        
        // contact two rows
        while (cells.hasNext()){
            Cell cell = cells.next();


            row.createCell(numberOfColumns++).setCellValue(cell.getStringCellValue());
        }

        System.out.print(row.getCell(0).getStringCellValue()+"--");
        System.out.print(row.getCell(EXP_Field.torder.ordinal()));
    }

    private void removeBackSlash(Row row) {
        for (Cell cell: row) {
            String field = cell.getStringCellValue();
            field = field.replace("\\", "");
            cell.setCellValue(field);
        }
    }

    public static void main(String[] args) throws IOException {
        InputStream FILE = new FileInputStream("src/AH_EXP_defualt.xlsx");
        XSSFWorkbook EXP = new XSSFWorkbook(FILE);

        Sheet sheet = EXP.getSheet("teacher_experience");

        Reorganize reorganizer = new Reorganize();
        reorganizer.reorganize(sheet);

        FILE.close();

        FileOutputStream output  = new FileOutputStream(new File("src/AH_EXP_reorganize.xlsx"));
        EXP.write(output);
        output.close();

    }
}
