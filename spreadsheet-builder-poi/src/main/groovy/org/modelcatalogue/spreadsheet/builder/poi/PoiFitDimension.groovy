package org.modelcatalogue.spreadsheet.builder.poi

import org.modelcatalogue.spreadsheet.api.Keywords.Fit
import org.modelcatalogue.spreadsheet.builder.api.FitDimension

class PoiFitDimension implements FitDimension{

    private final PoiPageSettingsProvider pageSettingsProvider
    private final Fit fit

    PoiFitDimension(PoiPageSettingsProvider pageSettingsProvider, Fit fit) {
        this.pageSettingsProvider = pageSettingsProvider
        this.fit = fit
    }

    @Override
    PoiPageSettingsProvider to(int numberOfPages) {
        switch (fit) {
            case Fit.HEIGHT:
                pageSettingsProvider.printSetup.setFitHeight(numberOfPages.shortValue())
                break
            case Fit.WIDTH:
                pageSettingsProvider.printSetup.setFitWidth(numberOfPages.shortValue())
                break
        }
        return pageSettingsProvider
    }
}
