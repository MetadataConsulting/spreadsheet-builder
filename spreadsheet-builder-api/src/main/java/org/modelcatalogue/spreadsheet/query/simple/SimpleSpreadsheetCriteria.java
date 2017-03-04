package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.query.api.*;

import java.io.FileNotFoundException;
import java.util.Iterator;

public final class SimpleSpreadsheetCriteria implements SpreadsheetCriteria {

    private final Workbook workbook;

    public static SpreadsheetCriteria forWorkbook(Workbook workbook) {
        return new SimpleSpreadsheetCriteria(workbook);
    }

    private SimpleSpreadsheetCriteria(Workbook workbook) {
        this.workbook = workbook;
    }

    private SpreadsheetCriteriaResult queryInternal(final int max, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") final Closure workbookCriterion) {
        return new SimpleSpreadsheetCriteriaResult(workbook, workbookCriterion, max);
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
