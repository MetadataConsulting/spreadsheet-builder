package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.PackageScope
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFName

@PackageScope class PendingFormula implements Resolvable {

    final XSSFCell cell
    final String formula

    PendingFormula(XSSFCell cell, String formula) {
        this.cell = cell
        this.formula = formula
    }

    void resolve() {
        cell.setCellFormula(expandNames(formula))
        cell.cellType = Cell.CELL_TYPE_FORMULA
    }

    protected String expandNames(String withNames) {
        withNames.replaceAll(/\#\{(.+?)\}/) { List<String> found ->
            XSSFName nameFound = cell.sheet.workbook.getName(PoiCell.fixName(found[1]))
            if (!found) {
                throw new IllegalArgumentException("Named cell '${found[1]}' cannot be found! Please, take a note that the name was normalized to ${PoiCell.fixName(found[1])} due the Excel constraints.")
            }
            nameFound.refersToFormula
        }
    }
}
