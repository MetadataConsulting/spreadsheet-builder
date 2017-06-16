package org.modelcatalogue.spreadsheet.builder.api.tck.java8;

import org.modelcatalogue.spreadsheet.builder.api.CanDefineStyle;
import org.modelcatalogue.spreadsheet.builder.api.Stylesheet;

import static org.modelcatalogue.spreadsheet.api.Keywords.*;
import static org.modelcatalogue.spreadsheet.api.Color.*;

class MyStyles implements Stylesheet {

    public void declareStyles(CanDefineStyle stylable) {
        stylable.style("h1", s -> {
            s.foreground(whiteSmoke)
            .fill(solidForeground)
            .font(f -> {
                f.size(22);
            });
        });
        stylable.style("h2", s -> {
            s.base("h1");
            s.font(f -> {
                f.size(16);
            });
        });
    }
}