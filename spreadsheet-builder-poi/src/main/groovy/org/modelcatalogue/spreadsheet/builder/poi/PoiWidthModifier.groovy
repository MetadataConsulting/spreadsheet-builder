package org.modelcatalogue.spreadsheet.builder.poi

import org.modelcatalogue.spreadsheet.builder.api.DimensionModifier

class PoiWidthModifier implements DimensionModifier {

    private PoiCellDefinition cell
    private final double width

    PoiWidthModifier(PoiCellDefinition cell, double width) {
        this.cell = cell
        this.width = width
    }

    @Override
    PoiCellDefinition cm() {
        cell.width(width * 4.6666666666666666666667)
        return cell
    }

    @Override
    PoiCellDefinition inch() {
        cell.width(width * 12)
        return cell
    }

    @Override
    PoiCellDefinition inches() {
        return inch
    }

    @Override
    PoiCellDefinition points() {
        return cell
    }
}
