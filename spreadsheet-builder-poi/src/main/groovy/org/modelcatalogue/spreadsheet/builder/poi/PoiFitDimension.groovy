package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.xssf.usermodel.XSSFPrintSetup
import org.modelcatalogue.spreadsheet.api.Keywords.Fit
import org.modelcatalogue.spreadsheet.builder.api.FitDimension

class PoiFitDimension implements FitDimension{

    private final XSSFPrintSetup printSetup
    private final Fit fit

    PoiFitDimension(XSSFPrintSetup printSetup, Fit fit) {
        this.printSetup = printSetup
        this.fit = fit
    }

    @Override
    void to(int numberOfPages) {
        switch (fit) {
            case Fit.HEIGHT:
                printSetup.setFitHeight(numberOfPages.shortValue())
                break
            case Fit.WIDTH:
                printSetup.setFitWidth(numberOfPages.shortValue())
                break
        }
    }
}
