package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Row;
import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.api.Workbook;
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetQuery;
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;

public final class SimpleSpreadsheetQuery implements SpreadsheetQuery {

    private final WorkbookSupplier loader;

    public SimpleSpreadsheetQuery(WorkbookSupplier loader) {
        this.loader = loader;
    }

    private Collection<Cell> queryInternal(InputStream inputStream, int max, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) {
        Workbook poiWorkbook = loader.load(inputStream);
        Collection<Cell> cells = new ArrayList<Cell>();
        SimpleWorkbookCriterion criterion = new SimpleWorkbookCriterion();
        DefaultGroovyMethods.with(criterion, workbookCriterion);

        for (Sheet sheet : poiWorkbook.getSheets()) {
            if (criterion.passesAnyCondition(sheet)) {
                for (Row row : sheet.getRows()) {
                    if (criterion.getCriteria().isEmpty()) {
                        cells.addAll(row.getCells());
                        if (cells.size() >= max) {
                            return cells;
                        }
                    } else {
                        for (SimpleSheetCriterion sheetCriterion : criterion.getCriteria()) {
                            if (sheetCriterion.passesAnyCondition(row)) {
                                if (sheetCriterion.getCriteria().isEmpty()) {
                                    cells.addAll(row.getCells());
                                    if (cells.size() >= max) {
                                        return cells;
                                    }
                                } else {
                                    for (Cell cell : row.getCells()) {
                                        for (SimpleRowCriterion rowCriterion : sheetCriterion.getCriteria()) {
                                            if (rowCriterion.passesAnyCondition(cell)) {
                                                if (rowCriterion.getCriteria().isEmpty()) {
                                                    cells.add(cell);
                                                    if (cells.size() >= max) {
                                                        return cells;
                                                    }
                                                } else {
                                                    for (SimpleCellCriterion cellCriterion : rowCriterion.getCriteria()) {
                                                        if (!cellCriterion.passesAllConditions(cell)) {
                                                            continue;
                                                        }
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
            }
        }

        return cells;
    }


    @Override
    public Collection<Cell> query(File spreadsheet) throws FileNotFoundException {
        return query(spreadsheet, Closure.IDENTITY);
    }

    @Override
    public Collection<Cell> query(File spreadsheet, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        return query(new FileInputStream(spreadsheet), workbookCriterion);
    }

    @Override
    public Collection<Cell> query(InputStream inputStream) {
        return query(inputStream, Closure.IDENTITY);
    }

    @Override
    public Collection<Cell> query(InputStream inputStream, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) {
        return queryInternal(inputStream, Integer.MAX_VALUE, workbookCriterion);
    }

    @Override
    public Cell find(File spreadsheet, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        return find(new FileInputStream(spreadsheet), workbookCriterion);
    }

    @Override
    public Cell find(InputStream inputStream, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) {
        Collection<Cell> cells = queryInternal(inputStream, 1, workbookCriterion);
        if (cells.size() > 0) {
            return cells.iterator().next();
        }
        return null;
    }

    @Override
    public boolean exists(File spreadsheet, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        return find(spreadsheet, workbookCriterion) != null;
    }

    @Override
    public boolean exists(InputStream inputStream, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) {
        return find(inputStream, workbookCriterion) != null;
    }
}
