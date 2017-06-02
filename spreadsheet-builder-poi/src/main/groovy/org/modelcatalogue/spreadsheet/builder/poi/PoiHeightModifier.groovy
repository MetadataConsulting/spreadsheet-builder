package org.modelcatalogue.spreadsheet.builder.poi

import org.modelcatalogue.spreadsheet.builder.api.CellDefinition
import org.modelcatalogue.spreadsheet.builder.api.DimensionModifier

class PoiHeightModifier implements DimensionModifier {

    private final PoiCellDefinition cell
    private final double height

    PoiHeightModifier(PoiCellDefinition cell, double height) {
        this.height = height
        this.cell = cell
    }

    @Override
    PoiCellDefinition getCm() {
        cell.height(28 * height)
        return cell
    }

    @Override
    PoiCellDefinition getInch() {
        cell.height(72 * height)
        return cell
    }

    @Override
    PoiCellDefinition getInches() {
        return inch
    }

    @Override
    CellDefinition getPoints() {
        return cell
    }
}
