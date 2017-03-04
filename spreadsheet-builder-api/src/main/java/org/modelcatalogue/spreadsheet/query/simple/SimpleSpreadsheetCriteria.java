package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.query.api.*;

import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.LinkedHashSet;

public final class SimpleSpreadsheetCriteria implements SpreadsheetCriteria {

    private final Workbook workbook;

    public static SpreadsheetCriteria forWorkbook(Workbook workbook) {
        return new SimpleSpreadsheetCriteria(workbook);
    }

    private SimpleSpreadsheetCriteria(Workbook workbook) {
        this.workbook = workbook;
    }

    private SpreadsheetCriteriaResult queryInternal(final int max, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") final Closure workbookCriterion) {

        return new AbstractSpreadsheetCriteriaResult() {
            @Override
            public Collection<Cell> getCells() {
                Collection<Cell> cells = new LinkedHashSet<Cell>();
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
            public Collection<Row> getRows() {
                Collection<Row> rows = new LinkedHashSet<Row>();
                SimpleWorkbookCriterion criterion = new SimpleWorkbookCriterion();
                DefaultGroovyMethods.with(criterion, workbookCriterion);

                for (Sheet sheet : workbook.getSheets()) {
                    if (criterion.test(sheet)) {
                        rows_loop:
                        for (Row row : sheet.getRows()) {
                            if (criterion.getCriteria().isEmpty()) {
                                rows.add(row);
                                if (rows.size() >= max) {
                                    return rows;
                                }
                            } else {
                                for (SimpleSheetCriterion sheetCriterion : criterion.getCriteria()) {
                                    if (sheetCriterion.test(row)) {
                                        if (sheetCriterion.getCriteria().isEmpty()) {
                                            rows.add(row);
                                            if (rows.size() >= max) {
                                                return rows;
                                            }
                                        } else {
                                            for (Cell cell : row.getCells()) {
                                                for (SimpleRowCriterion rowCriterion : sheetCriterion.getCriteria()) {
                                                    if (rowCriterion.test(cell)) {
                                                        rows.add(row);
                                                        if (rows.size() >= max) {
                                                            return rows;
                                                        }
                                                        continue rows_loop;
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

                return rows;
            }

            @Override
            public Collection<Sheet> getSheets() {
                Collection<Sheet> sheets = new LinkedHashSet<Sheet>();
                SimpleWorkbookCriterion criterion = new SimpleWorkbookCriterion();
                DefaultGroovyMethods.with(criterion, workbookCriterion);

                sheets_loop:
                for (Sheet sheet : workbook.getSheets()) {
                    if (criterion.test(sheet)) {
                        for (Row row : sheet.getRows()) {
                            if (criterion.getCriteria().isEmpty()) {
                                sheets.add(sheet);
                                if (sheets.size() >= max) {
                                    return sheets;
                                }
                            } else {
                                for (SimpleSheetCriterion sheetCriterion : criterion.getCriteria()) {
                                    if (sheetCriterion.test(row)) {
                                        if (sheetCriterion.getCriteria().isEmpty()) {
                                            sheets.add(sheet);
                                            if (sheets.size() >= max) {
                                                return sheets;
                                            }
                                        } else {
                                            for (Cell cell : row.getCells()) {
                                                for (SimpleRowCriterion rowCriterion : sheetCriterion.getCriteria()) {
                                                    if (rowCriterion.test(cell)) {
                                                        sheets.add(sheet);
                                                        if (sheets.size() >= max) {
                                                            return sheets;
                                                        }
                                                        continue sheets_loop;
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

                return sheets;
            }

        };


    }

    @Override
    public SpreadsheetCriteriaResult all() {
        return queryInternal(Integer.MAX_VALUE, Closure.IDENTITY);
    }

    @Override
    public SpreadsheetCriteriaResult query(@DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        return queryInternal(Integer.MAX_VALUE, workbookCriterion);
    }

    @Override
    public Cell find(@DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        SpreadsheetCriteriaResult cells = queryInternal(1, workbookCriterion);
        Iterator<Cell> cellIterator = cells.iterator();
        if (cellIterator.hasNext()) {
            return cellIterator.next();
        }
        return null;
    }

    @Override
    public boolean exists(@DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        return find(workbookCriterion) != null;
    }

}
