package org.modelcatalogue.spreadsheet.api;

import java.util.List;

import static org.codehaus.groovy.runtime.DefaultGroovyMethods.reverse;
import static org.codehaus.groovy.runtime.DefaultGroovyMethods.toList;

public interface Cell {

    int getColumn();
    Object getValue();
    String getColumnAsString();
    <T> T read(Class<T> type);
    Row getRow();

    int getRowspan();
    int getColspan();

    String getName();
    Comment getComment();
    CellStyle getStyle();

    Cell above();
    Cell bellow();
    Cell left();
    Cell right();
    Cell aboveLeft();
    Cell aboveRight();
    Cell bellowLeft();
    Cell bellowRight();

    class Util {

        private Util() {}

        public static int parseColumn(String column) {
            char a = 'A';
            List<Character> chars = reverse(toList(column.toUpperCase().toCharArray()));
            int acc = 0;
            for (int i = chars.size() - 1 ; i >= 0; i--) {
                if (i == 0) {
                    acc += ((int) chars.get(i) - (int) a + 1);
                } else {
                    acc += 26 * i * ((int) chars.get(i) - (int) a + 1);
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
