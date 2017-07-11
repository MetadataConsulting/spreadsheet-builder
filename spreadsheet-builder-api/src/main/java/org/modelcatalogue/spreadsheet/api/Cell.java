package org.modelcatalogue.spreadsheet.api;

public interface Cell extends Spannable {

    int getColumn();
    Object getValue();
    String getColumnAsString();
    <T> T read(Class<T> type);
    Row getRow();

    String getName();
    Comment getComment();
    CellStyle getStyle();

    Cell getAbove();
    Cell getBellow();
    Cell getLeft();
    Cell getRight();
    Cell getAboveLeft();
    Cell getAboveRight();
    Cell getBellowLeft();
    Cell getBellowRight();

    class Util {

        private Util() {}

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

}
