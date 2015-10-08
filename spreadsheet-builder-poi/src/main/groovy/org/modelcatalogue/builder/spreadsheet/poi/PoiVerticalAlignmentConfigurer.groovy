package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.modelcatalogue.builder.spreadsheet.api.VerticalAlignmentConfigurer

@CompileStatic class PoiVerticalAlignmentConfigurer implements VerticalAlignmentConfigurer {

    final PoiCellStyle style

    PoiVerticalAlignmentConfigurer(PoiCellStyle style) {
        this.style = style
    }

    @Override
    Object getTop() {
        style.setVerticalAlignment(VerticalAlignment.TOP)
        return null
    }

    @Override
    Object getCenter() {
        style.setVerticalAlignment(VerticalAlignment.CENTER)
        return null
    }

    @Override
    Object getBottom() {
        style.setVerticalAlignment(VerticalAlignment.BOTTOM)
        return null
    }

    @Override
    Object getJustify() {
        style.setVerticalAlignment(VerticalAlignment.JUSTIFY)
        return null
    }
}
