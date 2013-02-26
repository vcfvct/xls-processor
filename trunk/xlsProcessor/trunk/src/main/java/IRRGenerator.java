import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created with IntelliJ IDEA. User: LiHa Date: 2/26/13
 */
public class IRRGenerator
{

    private Workbook wb;
    private String location;
    private int totalRowNumber;
    private Sheet dataSource;
    private File targetFile;
    private FileInputStream fis = null;

    public IRRGenerator()
    {

    }

    public IRRGenerator(File file)
    {
        location = file.getAbsolutePath();
        targetFile = file;
        wb = Util.getWorkBook(targetFile);
    }

    public void generate(JProgressBar bar, JTextArea info) throws IOException
    {
        Sheet datasource = wb.getSheetAt(0);
        String sourceSheetName = datasource.getSheetName();
        totalRowNumber = Util.getRealLastRowNumber(datasource);

        Util.removeSheetByName(wb, "return-result");
        Sheet resultSheet = wb.createSheet("return-result");

        bar.setValue(10);
        info.append("File load successfully!\n");
        bar.setValue(30);
        info.append("There are " + (totalRowNumber - 1) + " rows in this file. Start process...\n");
        info.append("Start process...\n");
        System.out.println("File load successfully!");
        System.out.println("There are " + (totalRowNumber - 1) + " rows in this file. Start process...");

        Cell fundCell = resultSheet.createRow(0).createCell(0);
        fundCell.setCellValue("Fund");
        Cell textCell = resultSheet.getRow(0).createCell(1);
        textCell.setCellValue("Net IRR");
        int startRow = 1;
        int fundCount = 0;

        while (startRow < totalRowNumber)
        {
            fundCount++;
            String fundName = datasource.getRow(startRow).getCell(0).getStringCellValue();
            //create name cell
            Cell nameCell = resultSheet.createRow(fundCount).createCell(0);
            nameCell.setCellValue(fundName);

            int endRow = getLastRowForFund(datasource, fundName, startRow);
            //excel start from 1.
            //poi start from 0
            int excelStartRow = startRow + 1;
            int excelEndRow = endRow + 1;
            //create IRR cell
            String sourceSheetNameWithSingleQuote = "'" + sourceSheetName + "'";
            String formula = "XIRR(" + sourceSheetNameWithSingleQuote + "!C" + excelStartRow + ":" + sourceSheetNameWithSingleQuote + "!C" + excelEndRow + "," +
                sourceSheetNameWithSingleQuote + "!B" + excelStartRow +
                ":" + sourceSheetNameWithSingleQuote + "!B" + excelEndRow + ")";
            Cell cell = resultSheet.getRow(fundCount).createCell(1);
            cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
            cell.setCellFormula(formula);

            startRow = endRow + 1;

            info.append(fundName + "'s IRR from line " + excelStartRow + " to row " + excelEndRow + " calculated...\n");
        }

        //fis.close();
        File file = new File(location);
        file.delete();
        try
        {
            FileOutputStream out = new FileOutputStream(location);
            wb.write(out);
            out.close();
        }
        catch (IOException up)
        {
            info.append("!!!!!!!!!!!!!!! FAILED. The process cannot access the file because it is being used by another process.\n");
            throw up;
        }

        info.append("******** Congratulations Linda!\n");
        info.append("File: " + targetFile.getName() + " process Complete. \n");
        info.append("The result is in the \"result-sheet\" of the original file.\n");
        bar.setValue(100);
        System.out.println("******** Congratulations Linda!");
        System.out.println("******** File: " + location + " process Complete. ");
        System.out.println("******** The result is in the \"result-sheet\" of the original file. ");
    }

    private int getLastRowForFund(Sheet sheet, String fundName, int startRow)
    {
        int endRow = startRow;
        for (int i = startRow; i <= totalRowNumber; i++)
        {
            if (sheet.getRow(i).getCell(0).getStringCellValue().equals(fundName))
            {
                endRow = i;
                continue;
            }

            break;
        }

        return endRow;
    }

    public static void main(String[] args)
    {
        try
        {
            String location = "c:\\Users\\liha\\Desktop\\testIRR.xlsx";
            FileInputStream fis = new FileInputStream(location);
            Workbook wb = new XSSFWorkbook(fis);
            Sheet datasource = wb.getSheetAt(0);
            int lastRowNumber = Util.getRealLastRowNumber(datasource) + 1;
            Cell cell = datasource.getRow(3).createCell(5);
            String formula = "XIRR(B2:B" + lastRowNumber + ",A2:A" + lastRowNumber + ")";
            cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
            cell.setCellFormula(formula);
            fis.close();
            File file = new File(location);
            file.delete();
            FileOutputStream out = new FileOutputStream(location);
            wb.write(out);
            out.close();

        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
    }

}
