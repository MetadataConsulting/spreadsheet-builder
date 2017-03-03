package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.query.api.Predicate;
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteria;
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion;

import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Collection;

public final class SimpleSpreadsheetCriteria implements SpreadsheetCriteria {

    private final Workbook workbook;

    public static SpreadsheetCriteria forWorkbook(Workbook workbook) {
        return new SimpleSpreadsheetCriteria(workbook);
    }

    private SimpleSpreadsheetCriteria(Workbook workbook) {
        this.workbook = workbook;
    }

    private Collection<Cell> queryInternal(int max, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) {

        Collection<Cell> cells = new ArrayList<Cell>();
        SimpleWorkbookCriterion criterion = new SimpleWorkbookCriterion();
        DefaultGroovyMethods.with(criterion, workbookCriterion);

        for (Sheet sheet : workbook.getSheets()) {
            if (criterion.test(sheet)) {
                for (Row row : sheet.getRows()) {
                    if (criterion.getCriteria().isEmpty()) {
                        cells.addAll(row.getCells());
                        if (cells.size() >= max) {
                            return cells;
                        }
                    } else {
                        for (SimpleSheetCriterion sheetCriterion : criterion.getCriteria()) {
                            if (sheetCriterion.test(row)) {
                                if (sheetCriterion.getCriteria().isEmpty()) {
                                    cells.addAll(row.getCells());
                                    if (cells.size() >= max) {
                                        return cells;
                                    }
                                } else {
                                    for (Cell cell : row.getCells()) {
                                        for (SimpleRowCriterion rowCriterion : sheetCriterion.getCriteria()) {
                                            if (rowCriterion.test(cell)) {
                                                cells.add(cell);
                                                if (cells.size() >= max) {
                                                    return cells;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        return cells;
    }

    @Override
    public Collection<Cell> all() {
        return queryInternal(Integer.MAX_VALUE, Closure.IDENTITY);
    }

    @Override
    public Collection<Cell> query(@DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        return queryInternal(Integer.MAX_VALUE, workbookCriterion);
    }

    @Override
    public Cell find(@DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        Collection<Cell> cells = queryInternal(1, workbookCriterion);
        if (cells.size() > 0) {
            return cells.iterator().next();
        }
        return null;
    }

    @Override
    public boolean exists(@DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        return find(workbookCriterion) != null;
    }

}
