package com.sbyl;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Objects;

public class Main {

    static final String INPUT_FILE_PATH = "C:\\Users\\hanji\\Documents\\WeChat Files\\wxid_0xiky9xxszp622\\FileStorage\\File\\2025-04\\input.xlsx";
    static final String OUTPUT_FILE_PATH = "C:\\Users\\hanji\\Documents\\WeChat Files\\wxid_0xiky9xxszp622\\FileStorage\\File\\2025-04\\output.txt";

    static final String[] DEFAULT_PATH = {
            INPUT_FILE_PATH,
            OUTPUT_FILE_PATH
    };

    static final String PRINT_FORMAT =
            "{\n" +
            "    ID   : %s\n" +
            "    1    : %s\n" +
            "    2    : %s\n" +
            "    3    : %s\n" +
            "    4    : %s\n" +
            "    5    : %s\n" +
            "    6    : %s\n" +
            "    7    : %s\n" +
            "    8    : %s\n" +
            "    9    : %s\n" +
            "    10   : %s\n" +
            "    11   : %s\n" +
            "    12   : %s\n" +
            "    state: %s\n" +
            "}\n";

    /**
     * String: id.
     * ArrayList: data and state, contains 13 strings.
     * 0 - 11: data.
     * 12: data safety state, there are 3 states,
     * state -1: invalid data,
     * state 0: valid and safe data,
     * state 1: valid but unsafe data.
     */
    static HashMap<String, ArrayList<String>> dataMap;

    private static String[] getFilePath(String[] args) {
        String[] paths = new String[2];
        if (args == null || args.length != 2)
            return DEFAULT_PATH;

        String inputFilePath = args[0];
        String outputFilePath = args[1];
        if (Utils.isBlank(inputFilePath) || Utils.isBlank(outputFilePath))
            return DEFAULT_PATH;

        paths[0] = inputFilePath;
        paths[1] = outputFilePath;
        return paths;
    }

    /**
     * @param args
     * args[0]: file path of input file
     */
    public static void main(String[] args) {
        String[] filePath = getFilePath(args);
        String inputFilePath = filePath[0];
        String outputFilePath = filePath[1];
        System.out.printf("Input file path: %s", inputFilePath);
        System.out.printf("Output file path: %s", outputFilePath);

        File inputFile = new File(inputFilePath);

        dataMap = new HashMap<>();

        try (Workbook workbook = WorkbookFactory.create(inputFile)) {
            Sheet sheet = workbook.getSheetAt(0);

            System.out.print("Start handle excel file.");

            String currentSerialNumber = null;

            // Check rows and add data to map.
            for (Row row : sheet) {

                /*
                 * Check if the row is actually empty.
                 */
                boolean isEmptyRow = true;

                for (Cell cell : row) {
                    if (Utils.isNotBlank(cell.toString())) {
                        isEmptyRow = false;
                        break;
                    }
                }

                /*
                 * If the row is actually empty, reset the item.
                 * If current serial number is not null, the current item should be abnormal.
                 * Abnormal data should be -1.
                 */
                if (isEmptyRow) {
                    if (currentSerialNumber != null) {
                        ArrayList<String> data = new ArrayList<>();
                        for (int i = 0; i < 13; i++) {
                            data.add("-1");
                        }
                        dataMap.put(currentSerialNumber, data);
                    }
                    currentSerialNumber = null;
                    continue;
                }

                /*
                 * Check the cells in the row,
                 * if row contains special markings,
                 * then record the row as title row or data row.
                 */
                boolean isDataRow = false;
                int markPosition = -1;

                for (int i = 0; i < row.getLastCellNum(); i++) {
                    Cell cell = row.getCell(i);
                    if (cell == null)
                        continue;
                    String value = cell.toString();

                    if (Utils.isNotBlank(value)) {

                        /*
                         * If "统一编号：" exist, mark this row as a start row of current item,
                         * and then record the serial number.
                         * if value equals "统一编号：", it should be an abnormal data.
                         * Just ignore it.
                         */
                        if (value.contains("统一编号：")) {
                            String[] split = value.split("：");
                            if (split.length == 2) {
                                currentSerialNumber = split[1].replaceAll("\\s+", "");
                            } else {
                                currentSerialNumber = null;
                            }
                            break;
                        }

                        /*
                         * If "平均水位" exist, mark this row as a data row of current item.
                         */
                        if (currentSerialNumber != null && value.contains("平均水位")) {
                            isDataRow = true;
                            markPosition = i;
                            break;
                        }
                    }
                }

                /*
                 * Record data if current row is data row.
                 * If data row is broken, add -1 list.
                 * Then reset item.
                 */
                if (isDataRow) {
                    int state = 0;
                    ArrayList<String> data = new ArrayList<>();

                    // If marking position != 1, then data order might be abnormal.
                    if (markPosition != 1)
                        state = 1;

                    // Get data list and check list size.
                    for (Cell cell : row) {
                        String value = cell.toString();
                        if (!Utils.isDataCell(value))
                            continue;

                        ArrayList<String> dataList = Utils.getDataList(value);
                        if (dataList.isEmpty())
                            continue;

                        /*
                         * If dataList size > 1, means there are more than 1 number in the cell.
                         * We recognize the first number as target number, but this could be wrong.
                         * So mark the data as unsafe data.
                         */
                        if (dataList.size() != 1)
                            state = 1;

                        data.add(dataList.getFirst());
                    }

                    // If data size != 12, means the row is broken, add -1 list.
                    if (data.size() != 12) {
                        data.clear();
                        for (int i = 0; i < 13; i++) {
                            data.add("-1");
                        }
                    } else {
                        data.add(state + "");
                    }
                    dataMap.put(currentSerialNumber, data);

                    // Reset item.
                    currentSerialNumber = null;
                }
            }

            // Print result.
            System.out.print("Start print result.");
            int totalDataCount = dataMap.size();
            int validDataCount = 0;
            int safeDataCount = 0;
            StringBuilder builder = new StringBuilder();
            for (String key : dataMap.keySet()) {
                ArrayList<String> data = dataMap.get(key);
                String output = String.format(
                        PRINT_FORMAT,
                        key,
                        data.get(0),
                        data.get(1),
                        data.get(2),
                        data.get(3),
                        data.get(4),
                        data.get(5),
                        data.get(6),
                        data.get(7),
                        data.get(8),
                        data.get(9),
                        data.get(10),
                        data.get(11),
                        data.get(12));
                System.out.print(output);
                builder.append(output);
                if (!Objects.equals(data.get(12), "-1"))
                    validDataCount++;
                if (Objects.equals(data.get(12), "0"))
                    safeDataCount++;
            }
            System.out.printf("Total number of data: %d\n", totalDataCount);
            System.out.printf("Total number of valid data: %d\n", validDataCount);
            System.out.printf("Total number of safe data: %d\n", safeDataCount);
            System.out.printf("Percentage of valid data: %f%%\n", (validDataCount * 100f) / totalDataCount);
            System.out.printf("Percentage of safe data: %f%%\n", (safeDataCount * 100f) / totalDataCount);
            Utils.writeText(new File(outputFilePath), builder.toString());
        } catch (IOException e) {
            //throw new RuntimeException(e);
        }
    }
}