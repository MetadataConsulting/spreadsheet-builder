package builders.dsl.spreadsheet.query.poi;
import builders.dsl.spreadsheet.api.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.util.*;

class PoiRow implements Row {

    private final XSSFRow xssfRow;
    private final PoiSheet sheet;

    private Map<Integer, PoiCell> cells;

    PoiRow(PoiSheet sheet, XSSFRow xssfRow) {
        this.sheet = sheet;
        this.xssfRow = xssfRow;
    }

    public List<builders.dsl.spreadsheet.api.Cell> getCells() {
        if (cells == null) {
             cells = new LinkedHashMap<Integer, PoiCell>();
            for (Cell cell : xssfRow) {
                cells.put(cell.getColumnIndex() + 1, new PoiCell(this, (XSSFCell) cell));
            }
        }

        return Collections.unmodifiableList(new ArrayList<builders.dsl.spreadsheet.api.Cell>(cells.values()));
    }

    @Override
    public int getNumber() {
        return xssfRow.getRowNum() + 1;
    }

    @Override
    public PoiSheet getSheet() {
        return sheet;
    }

    protected XSSFRow getRow() {
        return xssfRow;
    }


    @Override
    public PoiRow getAbove(int howMany) {
        return aboveOrBellow(-howMany);
    }

    @Override
    public PoiRow getAbove() {
        return getAbove(1);
    }

    @Override
    public PoiRow getBellow(int howMany) {
        return aboveOrBellow(howMany);
    }

    @Override
    public PoiRow getBellow() {
        return getBellow(1);
    }

    builders.dsl.spreadsheet.api.Cell getCellByNumber(int oneBasedColumnNumber) {
        return cells.get(oneBasedColumnNumber);
    }

    private PoiRow aboveOrBellow(int howMany) {
        if (xssfRow.getRowNum() + howMany < 0 || xssfRow.getRowNum() + howMany >  xssfRow.getSheet().getLastRowNum()) {
            return null;
        }
        PoiRow existing = sheet.getRowByNumber(getNumber() + howMany);
        if (existing != null) {
            return existing;
        }
        return null;
    }
}
