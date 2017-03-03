package org.modelcatalogue.spreadsheet.api;

/**
 * Created by ladin on 03.03.17.
 */
public interface PageSettingsProvider {

    Keywords.Orientation getPortrait();
    Keywords.Orientation getLandscape();

    Keywords.Fit getWidth();
    Keywords.Fit getHeight();

    Keywords.To getTo();

    Keywords.Paper getLetter();
    Keywords.Paper getLetterSmall();
    Keywords.Paper getTabloid();
    Keywords.Paper getLedger();
    Keywords.Paper getLegal();
    Keywords.Paper getStatement();
    Keywords.Paper getExecutive();
    Keywords.Paper getA3();
    Keywords.Paper getA4();
    Keywords.Paper getA4Small();
    Keywords.Paper getA5();
    Keywords.Paper getB4();
    Keywords.Paper getB5();
    Keywords.Paper getFolio();
    Keywords.Paper getQuarto();
    Keywords.Paper getStandard10x14();
    Keywords.Paper getStandard11x17();
}
