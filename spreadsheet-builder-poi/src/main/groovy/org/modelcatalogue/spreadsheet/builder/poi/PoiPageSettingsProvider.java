package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.ss.usermodel.PaperSize;
import org.apache.poi.ss.usermodel.PrintOrientation;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.api.Page;
import org.modelcatalogue.spreadsheet.builder.api.FitDimension;
import org.modelcatalogue.spreadsheet.builder.api.PageDefinition;

class PoiPageSettingsProvider implements PageDefinition {

    PoiPageSettingsProvider(PoiSheetDefinition sheet) {
        this.printSetup = sheet.getSheet().getPrintSetup();
    }

    @Override
    public PoiPageSettingsProvider orientation(Keywords.Orientation orientation) {
        switch (orientation) {
            case PORTRAIT:
                printSetup.setOrientation(PrintOrientation.PORTRAIT);
                break;
            case LANDSCAPE:
                printSetup.setOrientation(PrintOrientation.LANDSCAPE);
                break;
        }
        return this;
    }

    @Override
    public PoiPageSettingsProvider paper(Keywords.Paper paper) {
        switch (paper) {
            case LETTER:
                printSetup.setPaperSize(PaperSize.LETTER_PAPER);
                break;
            case LETTER_SMALL:
                printSetup.setPaperSize(PaperSize.LETTER_SMALL_PAPER);
                break;
            case TABLOID:
                printSetup.setPaperSize(PaperSize.TABLOID_PAPER);
                break;
            case LEDGER:
                printSetup.setPaperSize(PaperSize.LEDGER_PAPER);
                break;
            case LEGAL:
                printSetup.setPaperSize(PaperSize.LEGAL_PAPER);
                break;
            case STATEMENT:
                printSetup.setPaperSize(PaperSize.STATEMENT_PAPER);
                break;
            case EXECUTIVE:
                printSetup.setPaperSize(PaperSize.EXECUTIVE_PAPER);
                break;
            case A3:
                printSetup.setPaperSize(PaperSize.A3_PAPER);
                break;
            case A4:
                printSetup.setPaperSize(PaperSize.A4_PAPER);
                break;
            case A4_SMALL:
                printSetup.setPaperSize(PaperSize.A4_SMALL_PAPER);
                break;
            case A5:
                printSetup.setPaperSize(PaperSize.A5_PAPER);
                break;
            case B4:
                printSetup.setPaperSize(PaperSize.B4_PAPER);
                break;
            case B5:
                printSetup.setPaperSize(PaperSize.B5_PAPER);
                break;
            case FOLIO:
                printSetup.setPaperSize(PaperSize.FOLIO_PAPER);
                break;
            case QUARTO:
                printSetup.setPaperSize(PaperSize.QUARTO_PAPER);
                break;
            case STANDARD_10_14:
                printSetup.setPaperSize(PaperSize.STANDARD_PAPER_10_14);
                break;
            case STANDARD_11_17:
                printSetup.setPaperSize(PaperSize.STANDARD_PAPER_11_17);
                break;
        }
        return this;
    }

    @Override
    public FitDimension fit(Keywords.Fit widthOrHeight) {
        return new PoiFitDimension(this, widthOrHeight);
    }

    final XSSFPrintSetup getPrintSetup() {
        return printSetup;
    }

    private final XSSFPrintSetup printSetup;
}
