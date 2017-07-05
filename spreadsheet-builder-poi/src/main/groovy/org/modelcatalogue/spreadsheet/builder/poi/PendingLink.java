package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.modelcatalogue.spreadsheet.builder.api.Resolvable;

/**
 * Pending link is a link which needs to be resolved at the end of the build when all the named references are known.
 */
class PendingLink implements Resolvable {

    PendingLink(XSSFCell cell, final String name) {
        if (!name.equals(PoiCellDefinition.fixName(name))) {
            throw new IllegalArgumentException("Cannot call cell \'" + name + "\' as this is invalid identifier in Excel. Suggestion: " + PoiCellDefinition.fixName(name));
        }

        this.cell = cell;
        this.name = name;
    }

    public void resolve() {
        XSSFWorkbook workbook = cell.getRow().getSheet().getWorkbook();
        XSSFName xssfName = workbook.getName(name);

        if (xssfName == null) {
            throw new IllegalArgumentException("Name " + name + " does not exist!");
        }


        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(Hyperlink.LINK_DOCUMENT);
        link.setAddress(xssfName.getRefersToFormula());

        cell.setHyperlink(link);
    }

    public final XSSFCell getCell() {
        return cell;
    }

    public final String getName() {
        return name;
    }

    private final XSSFCell cell;
    private final String name;
}
