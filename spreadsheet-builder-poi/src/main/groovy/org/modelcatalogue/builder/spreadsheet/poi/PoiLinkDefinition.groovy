package org.modelcatalogue.builder.spreadsheet.poi

import org.apache.poi.common.usermodel.Hyperlink
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFHyperlink
import org.apache.poi.xssf.usermodel.XSSFName
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.builder.spreadsheet.api.LinkDefinition

class PoiLinkDefinition implements LinkDefinition {

    private final XSSFCell cell
    private final PoiWorkbook workbook

    PoiLinkDefinition(PoiWorkbook workbook, XSSFCell xssfCell) {
        this.cell = xssfCell
        this.workbook = workbook
    }

    @Override
    void name(String name) {
        workbook.addPendingLink(name, cell)
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
