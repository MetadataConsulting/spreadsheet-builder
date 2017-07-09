package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.modelcatalogue.spreadsheet.builder.api.LinkDefinition;
import org.modelcatalogue.spreadsheet.impl.Utils;

import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class PoiLinkDefinition implements LinkDefinition {
    public PoiLinkDefinition(PoiWorkbookDefinition workbook, PoiCellDefinition cell) {
        this.cell = cell;
        this.workbook = workbook;
    }

    @Override
    public PoiCellDefinition name(String name) {
        workbook.addPendingLink(name, cell);
        return cell;
    }

    @Override
    public PoiCellDefinition email(String email) {
        XSSFWorkbook workbook = cell.getRow().getSheet().getWorkbook().asType(XSSFWorkbook.class);
        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(Hyperlink.LINK_EMAIL);
        link.setAddress("mailto:" + email);
        cell.getCell().setHyperlink(link);
        return cell;
    }

    @Override
    public PoiCellDefinition email(final Map<String, ?> parameters, String email) {
        try {
            XSSFWorkbook workbook = cell.getRow().getSheet().getWorkbook().asType(XSSFWorkbook.class);
            XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(Hyperlink.LINK_EMAIL);
            List<String> params = new ArrayList<String>();
            for (Map.Entry<String, ?> parameter: parameters.entrySet()) {
                if (parameter.getValue() != null) {
                    params.add(URLEncoder.encode(parameter.getKey(), "UTF-8") + "=" + URLEncoder.encode(parameter.getValue().toString(), "UTF-8"));
                } else {
                    params.add(URLEncoder.encode(parameter.getKey(), "UTF-8"));
                }
            }
            link.setAddress("mailto:" + email + "?" + Utils.join(params, "&"));
            cell.getCell().setHyperlink(link);
            return cell;
        } catch (UnsupportedEncodingException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public PoiCellDefinition url(String url) {
        XSSFWorkbook workbook = cell.getRow().getSheet().getWorkbook().asType(XSSFWorkbook.class);
        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(Hyperlink.LINK_URL);
        link.setAddress(url);
        cell.getCell().setHyperlink(link);
        return cell;
    }

    @Override
    public PoiCellDefinition file(String path) {
        XSSFWorkbook workbook = cell.getRow().getSheet().getWorkbook().asType(XSSFWorkbook.class);
        XSSFHyperlink link = workbook.getCreationHelper().createHyperlink(Hyperlink.LINK_FILE);
        link.setAddress(path);
        cell.getCell().setHyperlink(link);
        return cell;
    }

    private final PoiCellDefinition cell;
    private final PoiWorkbookDefinition workbook;
}
