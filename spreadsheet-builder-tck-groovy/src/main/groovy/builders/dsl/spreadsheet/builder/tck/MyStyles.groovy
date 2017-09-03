package builders.dsl.spreadsheet.builder.tck

import builders.dsl.spreadsheet.builder.api.CanDefineStyle
import builders.dsl.spreadsheet.builder.api.Stylesheet

import static builders.dsl.spreadsheet.api.Color.*

class MyStyles implements Stylesheet {

    void declareStyles(CanDefineStyle stylable) {
        stylable.style('h1') {
            foreground whiteSmoke
            fill solidForeground
            font {
                size 22
            }
        }
        stylable.style('h2') {
            base 'h1'
            font {
                size 16
            }
        }
    }
}