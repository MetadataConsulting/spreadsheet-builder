package org.modelcatalogue.spreadsheet.api

import org.modelcatalogue.spreadsheet.impl.Utils
import spock.lang.Specification
import spock.lang.Unroll

/**
 * Tests for Cell
 */
class CellUtilSpec extends Specification {

    @Unroll
    void "parse column #column to number #index"() {
        expect:
        Utils.parseColumn(column) == index
            Utils.toColumn(index) == column
        where:
        column  | index
        'A'     | 1
        'B'     | 2
        'Z'     | 26
        'AA'    | 27
        'AB'    | 28
        'DA'    | 105
    }

}
