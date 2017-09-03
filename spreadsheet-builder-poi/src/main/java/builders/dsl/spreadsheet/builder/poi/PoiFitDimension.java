package builders.dsl.spreadsheet.builder.poi;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.builder.api.FitDimension;

class PoiFitDimension implements FitDimension {

    PoiFitDimension(PoiPageSettingsProvider pageSettingsProvider, Keywords.Fit fit) {
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
