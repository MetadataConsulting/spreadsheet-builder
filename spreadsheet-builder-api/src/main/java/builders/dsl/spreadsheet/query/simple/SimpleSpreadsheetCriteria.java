package builders.dsl.spreadsheet.query.simple;

import builders.dsl.spreadsheet.api.Cell;
import builders.dsl.spreadsheet.api.Workbook;
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteria;
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteriaResult;
import builders.dsl.spreadsheet.query.api.WorkbookCriterion;

import java.io.FileNotFoundException;
import java.util.Iterator;
import java.util.function.Consumer;

public final class SimpleSpreadsheetCriteria implements SpreadsheetCriteria {

    private final Workbook workbook;

    public static SpreadsheetCriteria forWorkbook(Workbook workbook) {
        return new SimpleSpreadsheetCriteria(workbook);
    }

    private SimpleSpreadsheetCriteria(Workbook workbook) {
        this.workbook = workbook;
    }

    private SpreadsheetCriteriaResult queryInternal(final int max, Consumer<WorkbookCriterion>  workbookCriterion) {
        return new SimpleSpreadsheetCriteriaResult(workbook, workbookCriterion, max);
    }

    @Override
    public SpreadsheetCriteriaResult all() {
        return queryInternal(Integer.MAX_VALUE, (w) -> {});
    }

    @Override
    public SpreadsheetCriteriaResult query(Consumer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException {
        return queryInternal(Integer.MAX_VALUE, workbookCriterion);
    }

    @Override
    public Cell find(Consumer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException {
        SpreadsheetCriteriaResult cells = queryInternal(1, workbookCriterion);
        Iterator<Cell> cellIterator = cells.iterator();
        if (cellIterator.hasNext()) {
            return cellIterator.next();
        }
        return null;
    }

    @Override
    public boolean exists(Consumer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException {
        return find(workbookCriterion) != null;
    }


}
