package builders.dsl.spreadsheet.builder.api.tck.java8;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.builder.api.CanDefineStyle;
import builders.dsl.spreadsheet.builder.api.Stylesheet;

import static builders.dsl.spreadsheet.api.Color.*;

public class MyStyles implements Stylesheet {

    public void declareStyles(CanDefineStyle stylable) {
        stylable.style("h1", s -> {
            s.foreground(whiteSmoke)
            .fill(Keywords.solidForeground)
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