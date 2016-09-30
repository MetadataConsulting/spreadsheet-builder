package org.modelcatalogue.spreadsheet.builder.poi

import groovy.transform.PackageScope
import org.apache.poi.common.usermodel.Hyperlink
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFHyperlink
import org.apache.poi.xssf.usermodel.XSSFName
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * Pending link is a link which needs to be resolved at the end of the build when all the named references are known.
 */
@PackageScope class PendingLink implements Resolvable {

    final XSSFCell cell
    final String name

    PendingLink(XSSFCell cell, String name) {
        this.cell = cell
        this.name = name
    }

    void resolve() {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFName xssfName = workbook.getName(PoiCellDefinition.fixName(name)) as XSSFName

        if (!xssfName) {
            throw new IllegalArgumentException("Name $name does not exist! Please consider that the name was normalized to ${PoiCellDefinition.fixName(name)}")
        }

        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_DOCUMENT) as XSSFHyperlink
        link.address = xssfName.refersToFormula

        cell.hyperlink = link
    }
}
