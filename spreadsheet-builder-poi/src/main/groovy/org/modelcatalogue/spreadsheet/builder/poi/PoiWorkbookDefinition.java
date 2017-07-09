package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Workbook;
import org.modelcatalogue.spreadsheet.impl.AbstractWorkbookDefinition;
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition;
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition;

import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class PoiWorkbookDefinition extends AbstractWorkbookDefinition implements WorkbookDefinition, Workbook, SpreadsheetDefinition {

    private final XSSFWorkbook workbook;

    public PoiWorkbookDefinition(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    @Override
    protected PoiSheetDefinition createSheet(String name) {
        XSSFSheet sheet = workbook.getSheet(WorkbookUtil.createSafeSheetName(name));
        return new PoiSheetDefinition(this, sheet != null ? sheet : workbook.createSheet(WorkbookUtil.createSafeSheetName(name)));
    }

    @Override
    protected PoiCellStyleDefinition createCellStyle() {
        return new PoiCellStyleDefinition(this);
    }

    XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public List<PoiSheetDefinition> getSheets() {
        List<PoiSheetDefinition> sheets = new ArrayList<PoiSheetDefinition>();
        for (Sheet s : workbook) {
            sheets.add(new PoiSheetDefinition(this, (XSSFSheet)s));
        }
        return Collections.unmodifiableList(sheets);
    }

    @Override
    public void writeTo(OutputStream outputStream) {
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void writeTo(File file) {
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException ignored) {
                    // do nothing
                }
            }
        }
    }

    public <T> T asType(Class<T> type) {
        if (type.isInstance(workbook)) {
            return type.cast(workbook);
        }
        return DefaultGroovyMethods.asType(this, type);
    }

    protected void addPendingFormula(String formula, PoiCellDefinition cell) {
        toBeResolved.add(new PoiPendingFormula(cell, formula));
    }

    protected void addPendingLink(String ref, PoiCellDefinition cell) {
        toBeResolved.add(new PoiPendingLink(cell, ref));
    }

}
