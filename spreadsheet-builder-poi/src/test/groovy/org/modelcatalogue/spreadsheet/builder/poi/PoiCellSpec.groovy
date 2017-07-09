package org.modelcatalogue.spreadsheet.builder.poi

import org.modelcatalogue.spreadsheet.impl.Utils
import spock.lang.Specification
import spock.lang.Unroll

class PoiCellSpec extends Specification {

    @Unroll
    def "normalize name #name to #result"() {
        expect:
            Utils.fixName(name) == result

        where:
            name        | result
            'C'         | /_C/
            'c'         | /_c/
            'R'         | /_R/
            'r'         | /_r/
            '10'        | /_10/
            '(5)'       | /_5_/
            '2.4'       | /_2.4/
            'foo'       | /foo/
            'foo bar'   | /foo_bar/
    }

}
