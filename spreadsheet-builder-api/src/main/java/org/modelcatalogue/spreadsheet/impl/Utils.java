package org.modelcatalogue.spreadsheet.impl;

import java.util.regex.Pattern;

public class Utils {
    // from DGM
    public static String join(Iterable<String> array, String separator) {
        StringBuilder buffer = new StringBuilder();
        boolean first = true;

        if (separator == null) separator = "";

        for (String value : array) {
            if (first) {
                first = false;
            } else {
                buffer.append(separator);
            }
            buffer.append(value);
        }
        return buffer.toString();
    }

    public static String fixName(String name) {
        if (name == null || name.length() == 0) {
            throw new IllegalArgumentException("Name cannot be null or empty!");
        }

        if (name.startsWith("c") || name.startsWith("C") || name.startsWith("r") || name.startsWith("R")) {
            return "_" + name;
        }

        name = name.replaceAll( "[^.0-9a-zA-Z_]","_");
        if (!Pattern.compile("^[abd-qs-zABD-QS-Z_].*").matcher(name).matches()){
            return fixName("_" + name);
        }

        return name;
    }

    public static int parseColumn(String column) {
        char a = 'A';
        char[] chars = new StringBuilder(column).reverse().toString().toCharArray();
        int acc = 0;
        for (int i = chars.length - 1; i >= 0; i--) {
            if (i == 0) {
                acc += (int) chars[i] - (int) a + 1;
            } else {
                acc += 26 * i * ((int) chars[i] - (int) a + 1);
            }
        }
        return acc;
    }

    public static String toColumn(int number) {
        char a = 'A';

        int rest = number % 26;
        int times = number / 26;

        if (rest == 0 && times == 1) {
            return "Z";
        }

        if (times > 0) {
            return toColumn(times) + (char) (rest + a - 1);
        }

        return "" + (char) (rest + a - 1);
    }
}
