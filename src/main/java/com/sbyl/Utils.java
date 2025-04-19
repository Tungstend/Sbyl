package com.sbyl;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;

import static java.nio.charset.StandardCharsets.UTF_8;

public class Utils {

    public static boolean isBlank(String str) {
        return str == null || str.trim().isEmpty();
    }

    public static boolean isNotBlank(String str) {
        return !isBlank(str);
    }

    public static void writeText(File file, String text) throws IOException {
        writeText(file, text, UTF_8);
    }

    public static void writeText(File file, String text, Charset charset) throws IOException {
        writeBytes(file, text.getBytes(charset));
    }

    public static void writeBytes(File file, byte[] data) throws IOException {
        writeBytes(file.toPath(), data);
    }

    public static void writeBytes(Path file, byte[] data) throws IOException {
        Files.createDirectories(file.getParent());
        Files.write(file, data);
    }

    public static String doubleToString(double number) {
        BigDecimal bd = new BigDecimal(Double.toString(number));
        bd = bd.stripTrailingZeros();
        return bd.toPlainString();
    }

    public static String truncateTo12Digits(String numberStr) {
        if (numberStr == null)
            return null;

        if (!numberStr.matches("^\\d+$"))
            return numberStr;

        return numberStr.length() > 12
                ? numberStr.substring(0, 12)
                : numberStr;
    }

    public static String cleanString(String input) {
        if (input == null)
            return null;

        return input.replaceAll("[^0-9\\s.-]", "");
    }

    public static boolean isNumberCell(String str) {
        if (str == null || str.isEmpty())
            return false;

        if (!str.matches("^[0-9.\\- ]+$"))
            return false;

        return str.chars().anyMatch(c -> c != ' ');
    }

    public static String processString(String input) {
        if (input == null)
            return null;

        return input
                .replaceAll("\\s*\\.\\s*", ".")
                .replaceAll("-\\s+(?=\\d)", "-")
                .replaceAll("(\\d)\\s+(-)", "$1 $2")
                .replaceAll("\\s+", " ")
                .trim();
    }

    public static String getNumber(String str) {
        if (str == null)
            return null;

        String processedStr = str.replaceAll("\\s+", "");
        if (processedStr.isEmpty())
            return null;

        if (processedStr.matches("^[+-]?(\\d+\\.?\\d*|\\.\\d+)([eE][+-]?\\d+)?$"))
            return processedStr;

        return null;
    }

    public static ArrayList<String> getDataList(String input) {
        ArrayList<String> list = new ArrayList<>();
        String rawData = processString(input);
        String[] dataArray = rawData.split(" ");
        for (String s : dataArray) {
            if (getNumber(s.trim()) != null)
                list.add(getNumber(s.trim()));
        }
        return list;
    }
}
