import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created with IntelliJ IDEA. User: LiHa Date: 2/26/13
 */
public class Util
{
    /**
     * get real row count for sheet
     * @param sheet
     * @return
     */
    public static int getRealLastRowNumber(Sheet sheet)
    {
        int number = sheet.getLastRowNum();
        for (int i = number; i >= 0; i--)
        {
            if (sheet.getRow(i).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK)
            {
                number = i;
                break;
            }
        }
        return number;
    }

    /**
     * get column count for sheet
     * @param sheet
     * @return
     */
    public static int getRealColumnNumber(Sheet sheet)
    {
        Row firstRow = sheet.getRow(0);
        int number = firstRow.getLastCellNum();
        for (int i = number; i >= 0; i--)
        {
            if (sheet.getRow(i).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK)
            {
                number = i;
                break;
            }
        }
        return number;
    }

    public static Workbook getWorkBook(File file){
        Workbook wb = null;
        try
        {
            FileInputStream fis = new FileInputStream(file);

            if (file.getName().contains("xlsx"))
            {
                wb = new XSSFWorkbook(fis);
            }
            else
            {
                POIFSFileSystem fs = new POIFSFileSystem(fis);
                wb = new HSSFWorkbook(fs);
            }
        }
        catch (IOException e)
        {
            System.out.println("...................Loading file Error. ");
            System.out.println();
            e.printStackTrace();
        }
       return wb;
    }

    public static void removeSheetByName(Workbook wb, String name)
    {
        for (int i = wb.getNumberOfSheets() - 1; i >= 0; i--)
        {
            Sheet tmpSheet = wb.getSheetAt(i);
            if (tmpSheet.getSheetName().equals(name))
            {
                wb.removeSheetAt(i);
                return;
            }
        }
    }

    public static void copyRow( Workbook wb, Row sourceRow, Row newRow)
    {
        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++)
        {
            // Grab a copy of the old/new cell
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null)
            {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            CellStyle newCellStyle = wb.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            ;
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (newCell.getCellComment() != null)
            {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null)
            {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType())
            {
                case Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }
    }

}
