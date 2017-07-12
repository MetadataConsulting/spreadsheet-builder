package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder;
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition;
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition;

import java.io.*;

public enum PoiSpreadsheetBuilder implements SpreadsheetBuilder {
    INSTANCE;

    @Override
    public SpreadsheetDefinition build(Configurer<WorkbookDefinition> workbookDefinition) {
        return buildInternal(new XSSFWorkbook(), workbookDefinition);
    }

    private static PoiWorkbookDefinition buildInternal(XSSFWorkbook workbook, Configurer<WorkbookDefinition> workbookDefinition) {
        PoiWorkbookDefinition poiWorkbook = new PoiWorkbookDefinition(workbook);
        Configurer.Runner.doConfigure(workbookDefinition, poiWorkbook);
        poiWorkbook.resolve();

        return poiWorkbook;
    }

    @Override
    public SpreadsheetDefinition build(InputStream template, Configurer<WorkbookDefinition> workbookDefinition) {
        try {
            return buildInternal(new XSSFWorkbook(template), workbookDefinition);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                template.close();
            } catch (IOException ignored) {
                // do nothing
            }
        }
    }

    @Override
    public SpreadsheetDefinition build(File template, Configurer<WorkbookDefinition> workbookDefinition) {
        try {
            return build(new FileInputStream(template), workbookDefinition);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

}
