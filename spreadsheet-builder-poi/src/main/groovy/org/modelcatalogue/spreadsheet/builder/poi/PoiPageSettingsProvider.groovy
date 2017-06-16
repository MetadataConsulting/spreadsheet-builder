package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.ss.usermodel.PaperSize
import org.apache.poi.ss.usermodel.PrintOrientation
import org.apache.poi.xssf.usermodel.XSSFPrintSetup
import org.modelcatalogue.spreadsheet.api.Keywords
import org.modelcatalogue.spreadsheet.api.Page
import org.modelcatalogue.spreadsheet.builder.api.FitDimension
import org.modelcatalogue.spreadsheet.builder.api.PageDefinition

class PoiPageSettingsProvider implements Page, PageDefinition {

    final XSSFPrintSetup printSetup

    PoiPageSettingsProvider(PoiSheetDefinition sheet) {
        this.printSetup = sheet.sheet.printSetup
    }

    @Override
    PoiPageSettingsProvider orientation(Keywords.Orientation orientation) {
        switch (orientation) {
            case Keywords.Orientation.PORTRAIT:
                printSetup.setOrientation(PrintOrientation.PORTRAIT)
                break
            case Keywords.Orientation.LANDSCAPE:
                printSetup.setOrientation(PrintOrientation.LANDSCAPE)
                break
        }
        this
    }

    @Override
    PoiPageSettingsProvider paper(Keywords.Paper paper) {
        switch (paper) {
            case Keywords.Paper.LETTER:
                printSetup.setPaperSize(PaperSize.LETTER_PAPER)
                break
            case Keywords.Paper.LETTER_SMALL:
                printSetup.setPaperSize(PaperSize.LETTER_SMALL_PAPER)
                break
            case Keywords.Paper.TABLOID:
                printSetup.setPaperSize(PaperSize.TABLOID_PAPER)
                break
            case Keywords.Paper.LEDGER:
                printSetup.setPaperSize(PaperSize.LEDGER_PAPER)
                break
            case Keywords.Paper.LEGAL:
                printSetup.setPaperSize(PaperSize.LEGAL_PAPER)
                break
            case Keywords.Paper.STATEMENT:
                printSetup.setPaperSize(PaperSize.STATEMENT_PAPER)
                break
            case Keywords.Paper.EXECUTIVE:
                printSetup.setPaperSize(PaperSize.EXECUTIVE_PAPER)
                break
            case Keywords.Paper.A3:
                printSetup.setPaperSize(PaperSize.A3_PAPER)
                break
            case Keywords.Paper.A4:
                printSetup.setPaperSize(PaperSize.A4_PAPER)
                break
            case Keywords.Paper.A4_SMALL:
                printSetup.setPaperSize(PaperSize.A4_SMALL_PAPER)
                break
            case Keywords.Paper.A5:
                printSetup.setPaperSize(PaperSize.A5_PAPER)
                break
            case Keywords.Paper.B4:
                printSetup.setPaperSize(PaperSize.B4_PAPER)
                break
            case Keywords.Paper.B5:
                printSetup.setPaperSize(PaperSize.B5_PAPER)
                break
            case Keywords.Paper.FOLIO:
                printSetup.setPaperSize(PaperSize.FOLIO_PAPER)
                break
            case Keywords.Paper.QUARTO:
                printSetup.setPaperSize(PaperSize.QUARTO_PAPER)
                break
            case Keywords.Paper.STANDARD_10_14:
                printSetup.setPaperSize(PaperSize.STANDARD_PAPER_10_14)
                break
            case Keywords.Paper.STANDARD_11_17:
                printSetup.setPaperSize(PaperSize.STANDARD_PAPER_11_17)
                break
        }
        this
    }

    @Override
    FitDimension fit(Keywords.Fit widthOrHeight) {
        return new PoiFitDimension(this, widthOrHeight)
    }

    @Override
    Keywords.Orientation getOrientation() {
        switch(printSetup.orientation) {
            case PrintOrientation.DEFAULT:
                return null
            case PrintOrientation.PORTRAIT:
                return Keywords.Orientation.PORTRAIT
            case PrintOrientation.LANDSCAPE:
                return Keywords.Orientation.LANDSCAPE
        }
        return null
    }

    @Override
    Keywords.Paper getPaper() {
        switch (printSetup.paperSizeEnum) {
            case PaperSize.LETTER_PAPER:
                return Keywords.Paper.LETTER
            case PaperSize.LETTER_SMALL_PAPER:
                return Keywords.Paper.LETTER_SMALL
            case PaperSize.TABLOID_PAPER:
                return Keywords.Paper.TABLOID
            case PaperSize.LEDGER_PAPER:
                return Keywords.Paper.LEDGER
            case PaperSize.LEGAL_PAPER:
                return Keywords.Paper.LEGAL
            case PaperSize.STATEMENT_PAPER:
                return Keywords.Paper.STATEMENT
            case PaperSize.EXECUTIVE_PAPER:
                return Keywords.Paper.EXECUTIVE
            case PaperSize.A3_PAPER:
                return Keywords.Paper.A3
            case PaperSize.A4_PAPER:
                return Keywords.Paper.A4
            case PaperSize.A4_SMALL_PAPER:
                return Keywords.Paper.A4_SMALL
            case PaperSize.A5_PAPER:
                return Keywords.Paper.A5
            case PaperSize.B4_PAPER:
                return Keywords.Paper.B4
            case PaperSize.B5_PAPER:
                return Keywords.Paper.B5
            case PaperSize.FOLIO_PAPER:
                return Keywords.Paper.FOLIO
            case PaperSize.QUARTO_PAPER:
                return Keywords.Paper.QUARTO
            case PaperSize.STANDARD_PAPER_10_14:
                return Keywords.Paper.STANDARD_10_14
            case PaperSize.STANDARD_PAPER_11_17:
                return Keywords.Paper.STANDARD_11_17
        }
        return null
    }
}
