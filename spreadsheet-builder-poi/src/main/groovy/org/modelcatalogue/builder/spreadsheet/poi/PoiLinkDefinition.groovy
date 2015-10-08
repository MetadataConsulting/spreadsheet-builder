package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import org.apache.poi.common.usermodel.Hyperlink
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFHyperlink
import org.apache.poi.xssf.usermodel.XSSFName
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.builder.spreadsheet.api.LinkDefinition

@CompileStatic class PoiLinkDefinition implements LinkDefinition {

    private final XSSFCell cell

    PoiLinkDefinition(XSSFCell xssfCell) {
        this.cell = xssfCell
    }

    @Override
    void name(String name) {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFName xssfName = workbook.getName(name) as XSSFName

        if (!xssfName) {
            throw new IllegalArgumentException("Name $name does not exist!")
        }

        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_DOCUMENT) as XSSFHyperlink
        link.address = xssfName.refersToFormula

        cell.hyperlink = link

    }

    @Override
    void email(String email) {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_EMAIL) as XSSFHyperlink
        link.address = "mailto:$email"
        cell.hyperlink = link
    }

    @Override
    void email(Map<String, ?> parameters, String email) {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_EMAIL) as XSSFHyperlink
        link.address = "mailto:$email?${parameters.collect { String key, value -> "${URLEncoder.encode(key, 'UTF-8')}=${value ? URLEncoder.encode(value.toString(), 'UTF-8') : ''}"}.join('&')}"
        cell.hyperlink = link
    }

    @Override
    void url(String url) {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_URL) as XSSFHyperlink
        link.address = url
        cell.hyperlink = link
    }

    @Override
    void file(String path) {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_FILE) as XSSFHyperlink
        link.address = path
        cell.hyperlink = link
    }
}
