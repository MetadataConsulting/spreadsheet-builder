package org.modelcatalogue.spreadsheet.builder.poi

import org.modelcatalogue.spreadsheet.builder.api.CanDefineStyle
import org.modelcatalogue.spreadsheet.builder.api.Stylesheet

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