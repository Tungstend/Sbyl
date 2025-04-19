package com.sbyl;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Objects;

public class Main {

    static final String INPUT_FILE_PATH = "C:\\Users\\hanji\\Documents\\WeChat Files\\wxid_0xiky9xxszp622\\FileStorage\\File\\2025-04\\2018\\";
    static final String OUTPUT_FILE_PATH = "C:\\Users\\hanji\\Documents\\WeChat Files\\wxid_0xiky9xxszp622\\FileStorage\\File\\2025-04\\2018\\output.xlsx";
    static final String LOG_FILE_PATH = "C:\\Users\\hanji\\Documents\\WeChat Files\\wxid_0xiky9xxszp622\\FileStorage\\File\\2025-04\\2018\\log.txt";

    static final String[] DEFAULT_PATH = {
            INPUT_FILE_PATH,
            OUTPUT_FILE_PATH,
            LOG_FILE_PATH
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
        String[] paths = new String[3];
        if (args == null || args.length != 3)
            return DEFAULT_PATH;

        String inputFilePath = args[0];
        String outputFilePath = args[1];
        String logFilePath = args[2];
        if (Utils.isBlank(inputFilePath) || Utils.isBlank(outputFilePath) || Utils.isBlank(logFilePath))
            return DEFAULT_PATH;

        paths[0] = inputFilePath;
        paths[1] = outputFilePath;
        paths[2] = logFilePath;
        return paths;
    }

    /**
     * @param args
     * args[0]: file path of input file
     * args[1]: file path of output file
     * args[2]: file path of log file
     */
    public static void main(String[] args) {
        String[] filePath = getFilePath(args);
        String inputFilePath = filePath[0];
        String outputFilePath = filePath[1];
        String logFilePath = filePath[2];
        System.out.printf("Input file path: %s\n", inputFilePath);
        System.out.printf("Output file path: %s\n", outputFilePath);
        System.out.printf("Log file path: %s\n", logFilePath);

        File inputFilePri = new File(inputFilePath + "input1.xlsx");

        dataMap = new HashMap<>();

        // Check rows and add data to map.
        System.out.print("Start handle excel file.\n");
        try (Workbook workbook = WorkbookFactory.create(inputFilePri)) {
            Sheet sheet = workbook.getSheetAt(0);

            String currentSerialNumber = null;

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
                                currentSerialNumber = split[1].replaceAll("\\D+", "");
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
                        state = 2;

                    // Get data list and check list size.
                    for (Cell cell : row) {
                        String value = cell.toString();
                        if (!Utils.isNumberCell(value))
                            continue;

                        ArrayList<String> dataList = Utils.getDataList(value);
                        if (dataList.isEmpty())
                            continue;

                        /*
                         * If dataList size > 1, means there are more than 1 number in the cell.
                         * We recognize the first number as target number, but this could be wrong.
                         * So mark the data as unsafe data.
                         */
                        if (dataList.size() != 1 && (state == 0 || state == 2))
                            state++;

                        data.add(dataList.getFirst());
                    }

                    // If data size != 12, means the row is broken, add -1 list.
                    if (data.size() != 12) {
                        data.clear();
                        for (int i = 0; i < 13; i++) {
                            data.add("-1");
                        }
                    } else {
                        data.add(Integer.toString(state));
                    }
                    dataMap.put(currentSerialNumber, data);

                    // Reset item.
                    currentSerialNumber = null;
                }
            }
        } catch (IOException e) {
            //throw new RuntimeException(e);
        }

        {
            // Print result.
            System.out.print("Start print result.\n");
            int totalDataCount = dataMap.size();
            int validDataCount = 0;
            int safeDataCount = 0;
            int level1DataCount = 0;
            int level2DataCount = 0;
            int level3DataCount = 0;
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
                System.out.print(output + "\n");
                builder.append(output);
                if (!Objects.equals(data.get(12), "-1"))
                    validDataCount++;
                if (Objects.equals(data.get(12), "0"))
                    safeDataCount++;
                if (Objects.equals(data.get(12), "1"))
                    level1DataCount++;
                if (Objects.equals(data.get(12), "2"))
                    level2DataCount++;
                if (Objects.equals(data.get(12), "3"))
                    level3DataCount++;
            }
            System.out.printf("Total number of data from input file: %d\n", totalDataCount);
            System.out.printf("Total number of valid data from input file: %d\n", validDataCount);
            System.out.printf("Total number of safe data from input file: %d\n", safeDataCount);
            System.out.printf("Total number of level 1 data from input file: %d\n", level1DataCount);
            System.out.printf("Total number of level 2 data from input file: %d\n", level2DataCount);
            System.out.printf("Total number of level 3 data from input file: %d\n", level3DataCount);
            System.out.printf("Percentage of valid data from input file: %f%%\n", (validDataCount * 100f) / totalDataCount);
            System.out.printf("Percentage of safe data from input file: %f%%\n", (safeDataCount * 100f) / totalDataCount);
            System.out.printf("Percentage of level 1 data from input file: %f%%\n", (level1DataCount * 100f) / totalDataCount);
            System.out.printf("Percentage of level 2 data from input file: %f%%\n", (level2DataCount * 100f) / totalDataCount);
            System.out.printf("Percentage of level 3 data from input file: %f%%\n", (level3DataCount * 100f) / totalDataCount);
            try {
                Utils.writeText(new File(logFilePath), builder.toString());
            } catch (IOException e) {
                //throw new RuntimeException(e);
            }
        }

        // Write result to output file
        System.out.print("Start writing result to excel file.\n");
        File inputFileSec = new File(inputFilePath + "input2.xlsx");
        File outputFile = new File(outputFilePath);
        try (Workbook workbook = WorkbookFactory.create(inputFileSec)) {
            Sheet sheet = workbook.getSheetAt(0);
            
            int totalDataCount = 0;
            int validDataCount = 0;
            int safeDataCount = 0;
            int level1DataCount = 0;
            int level2DataCount = 0;
            int level3DataCount = 0;

            for (Row row : sheet) {
                if (row.getCell(0).getCellType() != CellType.NUMERIC)
                    continue;

                totalDataCount++;
                
                // If data not found, write -1 and skip.
                ArrayList<String> dataList = dataMap.get(Double.toString(row.getCell(0).getNumericCellValue()).replace(".", "").replace("E11", ""));
                if (dataList == null || dataList.size() != 13 || dataList.getLast().equals("-1")) {
                    row.createCell(5).setCellValue(-1);
                    row.createCell(17).setCellValue(-1);
                    continue;
                }
                
                // Data found, write data and state.
                for (int i = 5; i < 18; i++) {
                    if (row.getCell(i) == null)
                        row.createCell(i);
                    row.getCell(i).setCellValue(Float.parseFloat(dataList.get(i - 5)));
                }
                validDataCount++;
                if (Objects.equals(dataList.get(12), "0"))
                    safeDataCount++;
                if (Objects.equals(dataList.get(12), "1"))
                    level1DataCount++;
                if (Objects.equals(dataList.get(12), "2"))
                    level2DataCount++;
                if (Objects.equals(dataList.get(12), "3"))
                    level3DataCount++;
            }

            // Save file.
            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                workbook.write(out);
                System.out.print("File written.\n");
            }
            
            // Print valid data percent.
            System.out.printf("Total number of data written to output file: %d\n", totalDataCount);
            System.out.printf("Total number of valid data written to output file: %d\n", validDataCount);
            System.out.printf("Total number of safe data written to output file: %d\n", safeDataCount);
            System.out.printf("Total number of level 1 data written to output file: %d\n", level1DataCount);
            System.out.printf("Total number of level 2 data written to output file: %d\n", level2DataCount);
            System.out.printf("Total number of level 3 data written to output file: %d\n", level3DataCount);
            System.out.printf("Percentage of valid data written to output file: %f%%\n", (validDataCount * 100f) / totalDataCount);
            System.out.printf("Percentage of safe data written to output file: %f%%\n", (safeDataCount * 100f) / totalDataCount);
            System.out.printf("Percentage of level 1 data written to output file: %f%%\n", (level1DataCount * 100f) / totalDataCount);
            System.out.printf("Percentage of level 2 data written to output file: %f%%\n", (level2DataCount * 100f) / totalDataCount);
            System.out.printf("Percentage of level 3 data written to output file: %f%%\n", (level3DataCount * 100f) / totalDataCount);
        } catch (IOException e) {
            //throw new RuntimeException(e);
        }
    }
}