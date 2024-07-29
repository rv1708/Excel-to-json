import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.util.Iterator;

public class ExcelToJsonConverter {

    public static void main(String[] args) {
        try {
            // Example JSON with byte array (in reality, you would get this from your JSON source)
            String jsonInput = "{\"excelBytes\":\"<your-base64-encoded-excel-file>\"}";

            // Parse the JSON input
            ObjectMapper objectMapper = new ObjectMapper();
            JsonNode jsonNode = objectMapper.readTree(jsonInput);
            String excelBase64 = jsonNode.get("excelBytes").asText();
            byte[] excelBytes = java.util.Base64.getDecoder().decode(excelBase64);

            // Convert the Excel bytes to InputStream
            ByteArrayInputStream bis = new ByteArrayInputStream(excelBytes);

            // Read the Excel workbook
            Workbook workbook = new XSSFWorkbook(bis);
            Sheet sheet1 = workbook.getSheetAt(0); // Get the first sheet (Sheet 1)

            // Create JSON array to hold the rows
            ArrayNode jsonArray = objectMapper.createArrayNode();

            // Iterate through rows
            Iterator<Row> rowIterator = sheet1.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                ObjectNode jsonObject = objectMapper.createObjectNode();

                // Iterate through cells
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String cellValue = getCellValue(cell);

                    jsonObject.put("Column" + cell.getColumnIndex(), cellValue);
                }

                jsonArray.add(jsonObject);
            }

            // Output the JSON array
            String jsonOutput = objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(jsonArray);
            System.out.println(jsonOutput);

            // Close the workbook and input stream
            workbook.close();
            bis.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Helper method to get cell value as string
    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return Double.toString(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
