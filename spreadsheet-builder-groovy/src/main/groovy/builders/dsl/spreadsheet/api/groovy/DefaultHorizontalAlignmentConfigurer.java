package builders.dsl.spreadsheet.api.groovy;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.builder.api.CellStyleDefinition;

class DefaultHorizontalAlignmentConfigurer implements HorizontalAlignmentConfigurer {
    DefaultHorizontalAlignmentConfigurer(CellStyleDefinition cellStyleDefinition, Keywords.VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        this.cellStyleDefinition = cellStyleDefinition;
    }

    @Override
    public CellStyleDefinition getRight() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.RIGHT);
    }

    @Override
    public CellStyleDefinition getLeft() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.LEFT);
    }

    @Override
    public CellStyleDefinition getGeneral() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.GENERAL);
    }

    @Override
    public CellStyleDefinition getCenter() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.CENTER);
    }

    @Override
    public CellStyleDefinition getFill() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.FILL);
    }

    @Override
    public CellStyleDefinition getJustify() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.JUSTIFY);
    }

    @Override
    public CellStyleDefinition getCenterSelection() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.CENTER_SELECTION);
    }

    private CellStyleDefinition cellStyleDefinition;
    private Keywords.VerticalAlignment verticalAlignment;
}
