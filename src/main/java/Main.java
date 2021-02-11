import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Iterator;

/**
 * @author Jamith Nimantha
 */
public class Main {

    /**
     * BuiltinFormats can be find in here
     * http://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BuiltinFormats.html
     */
    public static void main(String[] args) {
        try (FileInputStream file = new FileInputStream(new File("poc-percentage.xlsx"))) {
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheetAt = workbook.getSheetAt(0);
            Iterator<Row> iterator = sheetAt.iterator();
            iterator.next(); // skips the header
            while (iterator.hasNext()) {
                Row row = iterator.next();

                Cell agentCell = row.getCell(0); // get agent code
                Cell perCell = row.getCell(1); // get percentage

                String agentCode = agentCell.getStringCellValue();
                BigDecimal agentPercentage = getPercentage(perCell);

                System.out.printf("Agent : %s | Percentage : %s%%%n", agentCode, agentPercentage);
                System.out.println("====================================================");
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * @param cell percentage cell
     * @return true if the cell's data format is percentage(%)
     */
    private static boolean isPercentage(Cell cell) {
        return cell.getCellStyle().getDataFormatString().equals("0.00%");
    }

    /**
     * @param cell percentage cell
     * @return the percentage value of the numeric cell
     */
    private static BigDecimal getPercentage(Cell cell) {
        if (isPercentage(cell) || cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return BigDecimal.valueOf(cell.getNumericCellValue() * 100);
        }
        return BigDecimal.ZERO; // should handle an Exception
    }

}
