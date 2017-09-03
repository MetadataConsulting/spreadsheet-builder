package builders.dsl.spreadsheet.impl;

import builders.dsl.spreadsheet.builder.api.CellDefinition;
import builders.dsl.spreadsheet.builder.api.DimensionModifier;

public final class WidthModifier implements DimensionModifier {

    // FIXME: there's a potential risk in this API as calling e.g. .height(12).cm().cm() would power the width twice

    public WidthModifier(CellDefinition cell, double width, double pointsPerCentimeter, double pointsPerInch) {
        this.cell = cell;
        this.width = width;
        this.pointsPerCentimeter = pointsPerCentimeter;
        this.pointsPerInch = pointsPerInch;
    }

    @Override
    public CellDefinition cm() {
        cell.width(width * pointsPerCentimeter);
        return cell;
    }

    @Override
    public CellDefinition inch() {
        cell.width(width * pointsPerInch);
        return cell;
    }

    @Override
    public CellDefinition inches() {
        return inch();
    }

    @Override
    public CellDefinition points() {
        return cell;
    }

    private double pointsPerCentimeter;
    private double pointsPerInch;
    private CellDefinition cell;
    private final double width;
}
