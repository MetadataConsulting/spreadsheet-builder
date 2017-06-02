package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.common.usermodel.Hyperlink
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFHyperlink
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.builder.api.LinkDefinition

class PoiLinkDefinition implements LinkDefinition {

    private final PoiCellDefinition cell
    private final PoiWorkbookDefinition workbook

    PoiLinkDefinition(PoiWorkbookDefinition workbook, PoiCellDefinition cell) {
        this.cell = cell
        this.workbook = workbook
    }

    @Override
    PoiCellDefinition name(String name) {
        workbook.addPendingLink(name, cell.cell)
        return cell
    }

    @Override
    PoiCellDefinition email(String email) {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_EMAIL) as XSSFHyperlink
        link.address = "mailto:$email"
        cell.cell.hyperlink = link
        return cell
    }

    @Override
    PoiCellDefinition email(Map<String, ?> parameters, String email) {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_EMAIL) as XSSFHyperlink
        link.address = "mailto:$email?${parameters.collect { String key, value -> "${URLEncoder.encode(key, 'UTF-8')}=${value ? URLEncoder.encode(value.toString(), 'UTF-8') : ''}"}.join('&')}"
        cell.cell.hyperlink = link
        return cell
    }

    @Override
    PoiCellDefinition url(String url) {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_URL) as XSSFHyperlink
        link.address = url
        cell.cell.hyperlink = link
        return cell
    }

    @Override
    PoiCellDefinition file(String path) {
        XSSFWorkbook workbook = cell.row.sheet.workbook as XSSFWorkbook
        XSSFHyperlink link = workbook.creationHelper.createHyperlink(Hyperlink.LINK_FILE) as XSSFHyperlink
        link.address = path
        cell.cell.hyperlink = link
        return cell
    }
}
