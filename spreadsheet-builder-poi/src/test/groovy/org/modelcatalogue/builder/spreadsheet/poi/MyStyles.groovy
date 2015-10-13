package org.modelcatalogue.builder.spreadsheet.poi

import org.modelcatalogue.builder.spreadsheet.api.CanDefineStyle
import org.modelcatalogue.builder.spreadsheet.api.Stylesheet

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