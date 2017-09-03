package builders.dsl.spreadsheet.api;

import java.util.Collection;

public interface Sheet {

    String getName();
    Workbook getWorkbook();

    Collection<? extends Row> getRows();

    Page getPage();

    Sheet getNext();
    Sheet getPrevious();

    boolean isLocked();
    boolean isHidden();
    boolean isVisible();
    boolean isVeryHidden();
}
