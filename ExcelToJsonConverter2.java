import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelToJsonConverter {

    public static void main(String[] args) throws IOException {
        String excelFilePath = "path/to/your/excel/file.xlsx";
        String[] headers = {"Header1", "Header2", "Header3"}; // Replace with your actual headers

        FileInputStream fileInputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        int headerRowIndex = findHeaderRowIndex(sheet, headers);
        if (headerRowIndex != -1) {
            JSONArray jsonArray = convertRowsToJSON(sheet, headerRowIndex, headers);
            System.out.println(jsonArray.toString(2)); // Pretty print JSON
        } else {
            System.out.println("Header row not found!");
        }

        workbook.close();
        fileInputStream.close();
    }

    private static int findHeaderRowIndex(Sheet sheet, String[] headers) {
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (isHeaderRow(row, headers)) {
                return row.getRowNum();
            }
        }
        return -1;
    }

    private static boolean isHeaderRow(Row row, String[] headers) {
        Iterator<Cell> cellIterator = row.cellIterator();
        int matchCount = 0;
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            for (String header : headers) {
                if (cell.getStringCellValue().equalsIgnoreCase(header)) {
                    matchCount++;
                    break;
                }
            }
        }
        return matchCount == headers.length;
    }

    private static JSONArray convertRowsToJSON(Sheet sheet, int headerRowIndex, String[] headers) {
        JSONArray jsonArray = new JSONArray();
        for (int i = headerRowIndex + 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                JSONObject jsonObject = new JSONObject();
                for (int j = 0; j < headers.length; j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        jsonObject.put(headers[j], getCellValue(cell));
                    }
                }
                jsonArray.put(jsonObject);
            }
        }
        return jsonArray;
    }

    private static Object getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
