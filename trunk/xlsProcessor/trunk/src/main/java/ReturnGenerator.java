import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created with IntelliJ IDEA. User: liha Date: 1/11/13
 */
public class ReturnGenerator
{
    public static void main(String[] args) throws IOException, InvalidFormatException
    {
        if (args.length != 1)
        {

            System.out.println("******** Please input correct parameter format. ");
            System.out.println("******** This program only take one parameter: Excel file path.");
            System.out.println("******** If the file path contains space, please add quote around it. ");
            System.out.println("******** Example: java -jar \"C:\\my Excel.xls\"");
        }
        ReturnGenerator rg = new ReturnGenerator(new File(args[0]));
        //rg.generate();

    }

    private Workbook wb;
    private String location;
    private int totalRowNumber;
    private Sheet dataSource;
    private File targetFile;
    private FileInputStream fis = null;

    //    public ReturnGenerator(String location)
    //    {
    //        this.location = location;
    //        try
    //        {
    //            fis = new FileInputStream(location);
    //
    //        if (location.contains("xlsx"))
    //        {
    //            wb = new XSSFWorkbook(fis);
    //        }
    //        else
    //        {
    //            POIFSFileSystem fs = new POIFSFileSystem(fis);
    //            wb = new HSSFWorkbook(fs);
    //        }
    //        }
    //        catch (IOException e)
    //        {
    //            System.out.println("...................Loading file Error. ");
    //            System.out.println();
    //            e.printStackTrace();
    //        }
    //    }

    public ReturnGenerator(File file)
    {
        location = file.getAbsolutePath();
        targetFile = file;
        init();
    }

    public void init()
    {
        try
        {
            fis = new FileInputStream(targetFile);

            if (targetFile.getName().contains("xlsx"))
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
    }

    public void generate(JProgressBar bar, JTextArea info) throws IOException
    {

        dataSource = wb.getSheetAt(0);

        totalRowNumber = getRealLastRowNumber(dataSource);

        bar.setValue(10);
        info.append("File load successfully!\n");
        bar.setValue(30);
        info.append("There are " + (totalRowNumber - 1) + " rows in this file. Start process...\n");
        info.append("Start process...\n");
        System.out.println("File load successfully!");
        System.out.println("There are " + (totalRowNumber - 1) + " rows in this file. Start process...");

        removeSheetByName("return-result");
        Sheet resultSheet = wb.createSheet("return-result");
        Row firstRow = dataSource.getRow(0);
        copyRow(firstRow, resultSheet.createRow(0));

        if (totalRowNumber > 2)
        {
            Row oneMonthRow = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
            createCellForFirstColumn(dataSource, oneMonthRow, "1 Mo.");
            generateReturnByPeriod(dataSource, oneMonthRow, 1);
            info.append("1 Mo. calculated...\n");
        }
        if (totalRowNumber > 3)
        {
            Row twoMonthRow = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
            createCellForFirstColumn(dataSource, twoMonthRow, "2 Mo.");
            generateReturnByPeriod(dataSource, twoMonthRow, 2);
            info.append("2 Mo. calculated...\n");
        }
        if (totalRowNumber > 4)
        {
            Row threeMonthRow = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
            createCellForFirstColumn(dataSource, threeMonthRow, "3 Mo.");
            generateReturnByPeriod(dataSource, threeMonthRow, 3);
            info.append("3 Mo. calculated...\n");
        }
        if (totalRowNumber > 13)
        {
            Row twelveMonthRow = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
            createCellForFirstColumn(dataSource, twelveMonthRow, "12 Mo.");
            generateReturnByPeriod(dataSource, twelveMonthRow, 12);
            info.append("12 Mo calculated...\n");
        }

        if (totalRowNumber > 37)
        {
            Row threeYearlyRow = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
            createCellForFirstColumn(dataSource, threeYearlyRow, "36 Mo.");
            generateReturnByPeriod(dataSource, threeYearlyRow, 36);
            info.append("36 Mo calculated...\n");
        }

        if (totalRowNumber > 61)
        {
            Row fiveYearlyRow = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
            createCellForFirstColumn(dataSource, fiveYearlyRow, "60 Mo.");
            generateReturnByPeriod(dataSource, fiveYearlyRow, 60);
            info.append("60 Mo calculated...\n");
        }
        if (totalRowNumber > 121)
        {
            Row tenYearlyRow = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
            createCellForFirstColumn(dataSource, tenYearlyRow, "120 Mo.");
            generateReturnByPeriod(dataSource, tenYearlyRow, 120);
            info.append("120 Mo calculated...\n");
        }

        int lastJul = getLastJulyRowNumber(dataSource);
        if(lastJul != -1){
            Row fiscalYearlyRow = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
            createCellForFirstColumn(dataSource, fiscalYearlyRow, "Fiscal YTD");
            generateReturnByPeriod(dataSource, fiscalYearlyRow, totalRowNumber-lastJul+1);
            info.append("Fiscal YTD calculated...\n");
        }
        int lastJan = getLastJanuaryRowNumber(dataSource);
        if(lastJan != -1){
            Row yearlyRow = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
            createCellForFirstColumn(dataSource, yearlyRow, "YTD");
            generateReturnByPeriod(dataSource, yearlyRow, totalRowNumber-lastJan+1);
            info.append("YTD calculated...\n");
        }

        Row sinceInception = resultSheet.createRow(resultSheet.getLastRowNum() + 1);
        createCellForFirstColumn(dataSource, sinceInception, "Since Inception");
        generateReturnByPeriod(dataSource, sinceInception, totalRowNumber - 1);
        info.append("Since Inception calculated...\n");

        bar.setValue(70);
        fis.close();
        File file = new File(location);
        file.delete();
        bar.setValue(80);
        FileOutputStream out = new FileOutputStream(location);
        wb.write(out);
        out.close();

        info.append("******** Congratulations Linda!\n");
        info.append("File: " + targetFile.getName() + " process Complete. \n");
        info.append("The result is in the \"result-sheet\" of the original file.\n");
        bar.setValue(100);
        System.out.println("******** Congratulations Linda!");
        System.out.println("******** File: " + location + " process Complete. ");
        System.out.println("******** The result is in the \"result-sheet\" of the original file. ");
    }

    private int getRealLastRowNumber(Sheet sheet)
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

    private int getColumnNumber(Sheet sheet)
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

    // Copy style from cell in row 3, column 1
    private CellStyle getDefaultCellStyle(Sheet dataSource)
    {

        CellStyle newCellStyle = wb.createCellStyle();
        newCellStyle.cloneStyleFrom(dataSource.getRow(2).getCell(1).getCellStyle());

        return newCellStyle;
    }

    private Cell createCellForFirstColumn(Sheet dataSource, Row row, String cellName)
    {
        Cell cell = row.createCell(0);
        cell.setCellStyle(getDefaultCellStyle(dataSource));
        cell.setCellValue(cellName);

        return cell;
    }

    private void removeSheetByName(String name)
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

    private void copyRow(Row sourceRow, Row newRow)
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


    private int getLastJulyRowNumber(Sheet dataSource)
    {
        int lastJul = -1;
        for (int i = totalRowNumber; i > 1; i--)
        {
            Date cellValue =  dataSource.getRow(i).getCell(0).getDateCellValue();
            SimpleDateFormat sdf = new SimpleDateFormat("MM");
            String month = sdf.format(cellValue);
            if (month.equals("07"))
            {
                return i;
            }
        }
        return lastJul;
    }
    private int getLastJanuaryRowNumber(Sheet dataSource)
    {
        int lastJan = -1;
        for (int i = totalRowNumber; i > 1; i--)
        {
            Date cellValue =  dataSource.getRow(i).getCell(0).getDateCellValue();
            SimpleDateFormat sdf = new SimpleDateFormat("MM");
            String month = sdf.format(cellValue);
            if (month.equals("01"))
            {
                return i;
            }
        }
        return lastJan;
    }


    /**
     * @param dataSource the working sheet
     * @param resultRow the row to be appended
     * @param period the time period to be calculated, calculate from the bottom to top.
     */
    private void generateReturnByPeriod(Sheet dataSource, Row resultRow, int period)
    {
        int lastColumnNumber = getColumnNumber(dataSource);


        for (int column = 1; column < lastColumnNumber; column++)
        {

            double result = 1;
            for (int row = 0; row < period; row++)
            {
                result *= (1 + dataSource.getRow(totalRowNumber - row).getCell(column).getNumericCellValue());
            }
            result -= 1;
            Cell cell = resultRow.createCell(column);
            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            cell.setCellStyle(getDefaultCellStyle(dataSource));
            cell.setCellValue(result);
        }
    }
}
