package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Row;
import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.api.Workbook;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;
import org.modelcatalogue.spreadsheet.query.api.AbstractSpreadsheetCriteriaResult;
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion;

import java.util.Collection;
import java.util.LinkedHashSet;

final class SimpleSpreadsheetCriteriaResult extends AbstractSpreadsheetCriteriaResult {

    private final Workbook workbook;
    private final Configurer<WorkbookCriterion> workbookCriterion;
    private final int max;

    SimpleSpreadsheetCriteriaResult(Workbook workbook, Configurer<WorkbookCriterion> workbookCriterion, int max) {
        this.workbook = workbook;
        this.workbookCriterion = workbookCriterion;
        this.max = max;
    }

    private Collection<Cell> getCellsInternal(int currentMax) {
        Collection<Cell> cells = new LinkedHashSet<Cell>();
        SimpleWorkbookCriterion criterion = new SimpleWorkbookCriterion();
        Configurer.Runner.doConfigure(workbookCriterion, criterion);

        for (Sheet sheet : workbook.getSheets()) {
            if (criterion.test(sheet)) {
                for (Row row : sheet.getRows()) {
                    if (criterion.getCriteria().isEmpty()) {
                        cells.addAll(row.getCells());
                        if (cells.size() >= currentMax) {
                            return cells;
                        }
                    } else {
                        for (SimpleSheetCriterion sheetCriterion : criterion.getCriteria()) {
                            if (sheetCriterion.test(row)) {
                                if (sheetCriterion.getCriteria().isEmpty()) {
                                    cells.addAll(row.getCells());
                                    if (cells.size() >= currentMax) {
                                        return cells;
                                    }
                                } else {
                                    for (Cell cell : row.getCells()) {
                                        for (SimpleRowCriterion rowCriterion : sheetCriterion.getCriteria()) {
                                            if (rowCriterion.test(cell)) {
                                                cells.add(cell);
                                                if (cells.size() >= currentMax) {
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

    private Collection<Row> getRowsInternal(int currentMax) {
        Collection<Row> rows = new LinkedHashSet<Row>();
        SimpleWorkbookCriterion criterion = new SimpleWorkbookCriterion();
        Configurer.Runner.doConfigure(workbookCriterion, criterion);

        for (Sheet sheet : workbook.getSheets()) {
            if (criterion.test(sheet)) {
                rows_loop:
                for (Row row : sheet.getRows()) {
                    if (criterion.getCriteria().isEmpty()) {
                        rows.add(row);
                        if (rows.size() >= currentMax) {
                            return rows;
                        }
                    } else {
                        for (SimpleSheetCriterion sheetCriterion : criterion.getCriteria()) {
                            if (sheetCriterion.test(row)) {
                                if (sheetCriterion.getCriteria().isEmpty()) {
                                    rows.add(row);
                                    if (rows.size() >= currentMax) {
                                        return rows;
                                    }
                                } else {
                                    for (Cell cell : row.getCells()) {
                                        for (SimpleRowCriterion rowCriterion : sheetCriterion.getCriteria()) {
                                            if (rowCriterion.test(cell)) {
                                                rows.add(row);
                                                if (rows.size() >= currentMax) {
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

    private Collection<Sheet> getSheetsInternal(int currentMax) {
        Collection<Sheet> sheets = new LinkedHashSet<Sheet>();
        SimpleWorkbookCriterion criterion = new SimpleWorkbookCriterion();
        Configurer.Runner.doConfigure(workbookCriterion, criterion);

        sheets_loop:
        for (Sheet sheet : workbook.getSheets()) {
            if (criterion.test(sheet)) {
                for (Row row : sheet.getRows()) {
                    if (criterion.getCriteria().isEmpty()) {
                        sheets.add(sheet);
                        if (sheets.size() >= currentMax) {
                            return sheets;
                        }
                    } else {
                        for (SimpleSheetCriterion sheetCriterion : criterion.getCriteria()) {
                            if (sheetCriterion.test(row)) {
                                if (sheetCriterion.getCriteria().isEmpty()) {
                                    sheets.add(sheet);
                                    if (sheets.size() >= currentMax) {
                                        return sheets;
                                    }
                                } else {
                                    for (Cell cell : row.getCells()) {
                                        for (SimpleRowCriterion rowCriterion : sheetCriterion.getCriteria()) {
                                            if (rowCriterion.test(cell)) {
                                                sheets.add(sheet);
                                                if (sheets.size() >= currentMax) {
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

    @Override
    public Collection<Cell> getCells() {
        return getCellsInternal(max);
    }

    @Override
    public Collection<Row> getRows() {
        return getRowsInternal(max);
    }

    @Override
    public Collection<Sheet> getSheets() {
        return getSheetsInternal(max);
    }

    @Override
    public Cell getCell() {
        Collection<Cell> cells = getCellsInternal(1);
        if (cells.size() > 0) {
            return cells.iterator().next();
        }
        return null;
    }

    @Override
    public Row getRow() {
        Collection<Row> rows = getRowsInternal(1);
        if (rows.size() > 0) {
            return rows.iterator().next();
        }
        return null;
    }

    @Override
    public Sheet getSheet() {
        Collection<Sheet> sheets = getSheetsInternal(1);
        if (sheets.size() > 0) {
            return sheets.iterator().next();
        }
        return null;
    }

}