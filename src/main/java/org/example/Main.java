import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        String xlsxFilePath = "/home/raj/IdeaProjects/DataReader/src/main/resources/excel-files/Benchmark.xlsx"; // Update the path to your XLSX file
        ObjectMapper objectMapper = new ObjectMapper();
        long recored_number= 10000000000001L;
        long increment_recored_number= 1L;

        try (FileInputStream fis = new FileInputStream(new File(xlsxFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(3); // Get the sheet (4th sheet)


            ObjectNode response_state= objectMapper.createObjectNode();
            response_state.put("response-code", 200);
            response_state.put("status", "succeed");

            ArrayNode response_results_state = objectMapper.createArrayNode();

            for (int i = 3; i <= sheet.getLastRowNum(); i++) { // Start from row 1 (skip headers)
                Row row = sheet.getRow(i);
                if (row == null) continue;

                if(getCellValueAsString(row.getCell(11)).equals("advance"))
                    continue;

                ObjectNode result_state = objectMapper.createObjectNode();
                result_state.put( "rule.name", getCellValueAsString(row.getCell(3)));
                result_state.put("rule.category", "Network");

                ObjectNode result_context_state= objectMapper.createObjectNode();
                result_context_state.put( "rule.check.category", getCellValueAsString(row.getCell(11)));
                result_context_state.put("rule.check.type", getCellValueAsString(row.getCell(10)));

                ObjectNode result_context_condition_state= objectMapper.createObjectNode();
                result_context_condition_state.put( "condition", getCellValueAsString(row.getCell(12)));
                result_context_condition_state.put("result.pattern", getCellValueAsString(row.getCell(13)));

                String occurrenceString = getCellValueAsString(row.getCell(14));
                if (occurrenceString.equals("any")) {
                    result_context_condition_state.put("occurrence", -1);
                }
                else if (occurrenceString.equals("")) {
                    result_context_condition_state.put("occurrence", "");
                }

                else {
                    result_context_condition_state.put("occurrence",1);
                }

                result_context_condition_state.put("operator", getCellValueAsString(row.getCell(15)).toUpperCase());

                ArrayNode result_context_conditions = objectMapper.createArrayNode();
                if (!isNodeEmpty(result_context_condition_state)) {
                    result_context_conditions.add(result_context_condition_state);
                }
                // addIfNotEmpty(result_context_conditions,result_context_condition_state);
                result_context_state.set("rule.conditions",result_context_conditions);


                result_state.set("rule.context", result_context_state);

                result_state.put("rule.auto.remediation", getCellValueAsString(row.getCell(8)));
                result_state.put("rule.description", getCellValueAsString(row.getCell(5)));
                result_state.put("rule.severity", getCellValueAsString(row.getCell(17)).toUpperCase());

                ArrayNode tags = objectMapper.createArrayNode();
                result_state.set("rule.tags", tags);
                result_state.put("rule.rationale", row.getCell(6).toString());
                result_state.put("rule.impact", getCellValueAsString(row.getCell(7)));
                result_state.put("rule.default.value", getCellValueAsString(row.getCell(35)));
                result_state.put("rule.references", getCellValueAsString(row.getCell(34)));
                result_state.put("rule.additional.information", getCellValueAsString(row.getCell(16)));

                ArrayNode result_control_state= objectMapper.createArrayNode();

                ObjectNode result_control_state_1 = objectMapper.createObjectNode();
                String[] control_states_1 = getControlState(18, row);
                result_control_state_1.put("rule.control.version", getCellValueAsString(row.getCell(22)).equals("0.0")?"":getCellValueAsString(row.getCell(22)));
                result_control_state_1.put("rule.control.name", control_states_1[0]==null?"":control_states_1[0]);
                result_control_state_1.put("rule.control.description", control_states_1[2]==null?"":control_states_1[2]);
                result_control_state_1.set("rule.control.ig",getcontrol_igs(row,25,26,27));



                ObjectNode result_control_state_2= objectMapper.createObjectNode();
                String[] control_states_2 = getControlState(19, row);
                result_control_state_2.put("rule.control.version", getCellValueAsString(row.getCell(28)).equals("0.0")?"":getCellValueAsString(row.getCell(28)));
                result_control_state_2.put("rule.control.name", control_states_2[0]==null?"":control_states_2[0]);
                result_control_state_2.put("rule.control.description", control_states_2[2]==null ?"":control_states_2[2]);
                result_control_state_2.set("rule.control.ig",getcontrol_igs(row,31,32,33));

                if (!isNodeEmpty(result_control_state_1)) {
                    result_control_state.add(result_control_state_1);
                }
                if (!isNodeEmpty(result_control_state_2)) {
                    result_control_state.add(result_control_state_2);
                }

                result_state.set("rule.controls", result_control_state);
                //result_state.put("profile", getCellValueAsString(row.getCell(3)));
                result_state.put("id", recored_number);
                recored_number+= increment_recored_number;

                response_results_state.add(result_state);
            }

            response_state.set("results",response_results_state);

            // Print JSON array
            System.out.println(objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(response_state));

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean isNodeEmpty(ObjectNode node) {
        Iterator<Map.Entry<String, JsonNode>> fields = node.fields();
        while (fields.hasNext()) {
            Map.Entry<String, JsonNode> entry = fields.next();
            if (!entry.getValue().asText().isEmpty()) {
                return false; // Found a non-empty value
            }
        }
        return true; // All values are empty
    }




    private static void addIfNotEmpty(ObjectNode node, String fieldName, String value) {
        if (value != null && !value.trim().isEmpty()) {
            node.put(fieldName, value);
        }
    }

    private static void addIfNotEmpty(ObjectNode node, String fieldName, ArrayNode value) {
        if (value != null && value.size() > 0) {
            node.set(fieldName, value); // Use set instead of put for complex types
        }
    }



    public static void addIfNotEmpty(ArrayNode node, ObjectNode fieldName) {
        if (!fieldName.isEmpty()) {
            node.add(fieldName);
        }
    }


    private static ArrayNode getcontrol_igs(Row row, int col1, int col2, int col3) {
        ObjectMapper objectMapper = new ObjectMapper();
        ArrayNode controls = objectMapper.createArrayNode();

        String cellValue1 = getCellValueAsString(row.getCell(col1));
        if (cellValue1 != null && !cellValue1.trim().isEmpty()) {
            controls.add("ig1");
        }

        String cellValue2 = getCellValueAsString(row.getCell(col2));
        if (cellValue2 != null && !cellValue2.trim().isEmpty()) {
            controls.add("ig2");
        }

        String cellValue3 = getCellValueAsString(row.getCell(col3));
        if (cellValue3 != null && !cellValue3.trim().isEmpty()) {
            controls.add("ig3");
        }

        return controls;
    }


    private static String[] getControlState(int control_cell_number, Row control_row_number) {
        String cellValue = getCellValueAsString(control_row_number.getCell(control_cell_number));
        ArrayList<String> control_states = new ArrayList<>();

        String[] parts = cellValue.split(" (?=TITLE|CONTROL|DESCRIPTION)");

        for (String part : parts) {
            if (part.startsWith("TITLE:")) {
                control_states.add(part.substring("TITLE:".length()).trim());
            } else if (part.startsWith("CONTROL:")) {
                control_states.add(part.substring("CONTROL:".length()).trim());
            } else if (part.startsWith("DESCRIPTION:")) {
                control_states.add(part.substring("DESCRIPTION:".length()).trim());
            }
        }

        return control_states.toArray(new String[3]);
    }


    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
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
            case BLANK:
                return "";
            default:
                return "UNKNOWN";
        }
    }
}

