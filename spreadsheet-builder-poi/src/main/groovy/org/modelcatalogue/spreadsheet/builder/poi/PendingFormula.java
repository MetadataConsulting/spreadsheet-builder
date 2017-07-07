package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.modelcatalogue.spreadsheet.builder.api.Resolvable;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Pending formula is a formula definition which needs to be resolved at the end of the build where all named references
 * are know.
 */
class PendingFormula implements Resolvable {
    PendingFormula(XSSFCell cell, String formula) {
        this.cell = cell;
        this.formula = formula;
    }

    public void resolve() {
        cell.setCellFormula(expandNames(formula));
        cell.setCellType(Cell.CELL_TYPE_FORMULA);
    }

    private String expandNames(String withNames) {
        final Matcher matcher = Pattern.compile("#\\{(.+?)\\}").matcher(withNames);
        if (matcher.find()) {
            final StringBuffer sb = new StringBuffer(withNames.length() + 16);
            do {
                String replacement = findRefersToFormula(matcher.group(1));
                matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
            } while (matcher.find());
            matcher.appendTail(sb);
            return sb.toString();
        }
        return withNames;
    }

    private String findRefersToFormula(final String name) {
        if (!name.equals(PoiCellDefinition.fixName(name))) {
            throw new IllegalArgumentException("Name " + name + " is not valid Excel name! Suggestion: " + PoiCellDefinition.fixName(name));
        }

        XSSFName nameFound = cell.getSheet().getWorkbook().getName(name);
        if (nameFound == null) {
            throw new IllegalArgumentException("Named cell \'" + name + "\' cannot be found!");
        }

        return nameFound.getRefersToFormula();
    }

    public final XSSFCell getCell() {
        return cell;
    }

    public final String getFormula() {
        return formula;
    }

    private final XSSFCell cell;
    private final String formula;
}
