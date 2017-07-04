package org.modelcatalogue.spreadsheet.builder.poi;

import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.builder.api.FitDimension;

public class PoiFitDimension implements FitDimension {

    public PoiFitDimension(PoiPageSettingsProvider pageSettingsProvider, Keywords.Fit fit) {
        this.pageSettingsProvider = pageSettingsProvider;
        this.fit = fit;
    }

    @Override
    public PoiPageSettingsProvider to(int numberOfPages) {
        if (Keywords.Fit.HEIGHT.equals(fit)) {
            pageSettingsProvider.getPrintSetup().setFitHeight((short) numberOfPages);
            return pageSettingsProvider;
        }
        if (Keywords.Fit.WIDTH.equals(fit)) {
            pageSettingsProvider.getPrintSetup().setFitWidth((short) numberOfPages);
            return pageSettingsProvider;
        }
        throw new IllegalStateException();
    }

    private final PoiPageSettingsProvider pageSettingsProvider;
    private final Keywords.Fit fit;
}
