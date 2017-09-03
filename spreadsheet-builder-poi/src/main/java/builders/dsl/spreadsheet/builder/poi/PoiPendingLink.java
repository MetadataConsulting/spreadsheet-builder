package builders.dsl.spreadsheet.builder.poi;

import builders.dsl.spreadsheet.impl.AbstractPendingLink;
import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Pending link is a link which needs to be resolved at the end of the build when all the named references are known.
 */
class PoiPendingLink extends AbstractPendingLink {

    PoiPendingLink(PoiCellDefinition cell, final String name) {
        super(cell, name);
    }

    public void resolve() {
        XSSFWorkbook workbook = getPoiCell().getCell().getRow().getSheet().getWorkbook();
        XSSFName xssfName = workbook.getName(getName());

        if (xssfName == null) {
            throw new IllegalArgumentException("Name " + getName() + " does not exist!");
        }


        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(Hyperlink.LINK_DOCUMENT);
        link.setAddress(xssfName.getRefersToFormula());

        getPoiCell().getCell().setHyperlink(link);
    }

    private PoiCellDefinition getPoiCell() {
        return (PoiCellDefinition) getCell();
    }

}
