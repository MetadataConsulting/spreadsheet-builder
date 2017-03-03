package org.modelcatalogue.spreadsheet.builder.poi

import org.modelcatalogue.spreadsheet.builder.api.DimensionModifier

class PoiHeightModifier implements DimensionModifier {

    private final PoiCellDefinition cell
    private final double height

    PoiHeightModifier(PoiCellDefinition cell, double height) {
        this.height = height
        this.cell = cell
    }

    @Override
    Object getCm() {
        cell.height(28 * height)
        return null
    }

    @Override
    Object getInch() {
        cell.height(72 * height)
        return null
    }

    @Override
    Object getInches() {
        inch
    }
}
