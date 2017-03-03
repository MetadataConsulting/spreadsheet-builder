package org.modelcatalogue.spreadsheet.api;

public abstract class AbstractPageSettingsProvider implements PageSettingsProvider {

    @Override
    public Keywords.Orientation getPortrait() {
        return Keywords.Orientation.PORTRAIT;
    }

    @Override
    public Keywords.Orientation getLandscape() {
        return Keywords.Orientation.LANDSCAPE;
    }

    @Override
    public Keywords.Fit getWidth() {
        return Keywords.Fit.WIDTH;
    }

    @Override
    public Keywords.Fit getHeight() {
        return Keywords.Fit.HEIGHT;
    }

    @Override
    public Keywords.To getTo() {
        return Keywords.To.TO;
    }

    @Override
    public Keywords.Paper getLetter() {
        return Keywords.Paper.LETTER;
    }

    @Override
    public Keywords.Paper getLetterSmall() {
        return Keywords.Paper.LETTER_SMALL;
    }

    @Override
    public Keywords.Paper getTabloid() {
        return Keywords.Paper.TABLOID;
    }

    @Override
    public Keywords.Paper getLedger() {
        return Keywords.Paper.LEDGER;
    }

    @Override
    public Keywords.Paper getLegal() {
        return Keywords.Paper.LEGAL;
    }

    @Override
    public Keywords.Paper getStatement() {
        return Keywords.Paper.STATEMENT;
    }

    @Override
    public Keywords.Paper getExecutive() {
        return Keywords.Paper.EXECUTIVE;
    }

    @Override
    public Keywords.Paper getA3() {
        return Keywords.Paper.A3;
    }

    @Override
    public Keywords.Paper getA4() {
        return Keywords.Paper.A4;
    }

    @Override
    public Keywords.Paper getA4Small() {
        return Keywords.Paper.A4_SMALL;
    }

    @Override
    public Keywords.Paper getA5() {
        return Keywords.Paper.A5;
    }

    @Override
    public Keywords.Paper getB4() {
        return Keywords.Paper.B4;
    }

    @Override
    public Keywords.Paper getB5() {
        return Keywords.Paper.B5;
    }

    @Override
    public Keywords.Paper getFolio() {
        return Keywords.Paper.FOLIO;
    }

    @Override
    public Keywords.Paper getQuarto() {
        return Keywords.Paper.QUARTO;
    }

    @Override
    public Keywords.Paper getStandard10x14() {
        return Keywords.Paper.STANDARD_10_14;
    }

    @Override
    public Keywords.Paper getStandard11x17() {
        return Keywords.Paper.STANDARD_11_17;
    }
}
