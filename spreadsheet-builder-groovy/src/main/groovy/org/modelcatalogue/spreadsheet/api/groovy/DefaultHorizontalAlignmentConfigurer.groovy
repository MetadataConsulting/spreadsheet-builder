package org.modelcatalogue.spreadsheet.api.groovy

import org.modelcatalogue.spreadsheet.api.Keywords
import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition

class DefaultHorizontalAlignmentConfigurer implements HorizontalAlignmentConfigurer {
    private CellStyleDefinition cellStyleDefinition
    private Keywords.VerticalAlignment verticalAlignment

    DefaultHorizontalAlignmentConfigurer(CellStyleDefinition cellStyleDefinition, Keywords.VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment
        this.cellStyleDefinition = cellStyleDefinition
    }

    @Override
    CellStyleDefinition getRight() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.RIGHT)
    }

    @Override
    CellStyleDefinition getLeft() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.LEFT)
    }

    @Override
    CellStyleDefinition getGeneral() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.GENERAL)
    }

    @Override
    CellStyleDefinition getCenter() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.CENTER)
    }

    @Override
    CellStyleDefinition getFill() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.FILL)
    }

    @Override
    CellStyleDefinition getJustify() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.JUSTIFY)
    }

    @Override
    CellStyleDefinition getCenterSelection() {
        return cellStyleDefinition.align(verticalAlignment, Keywords.HorizontalAlignment.CENTER_SELECTION)
    }
}
