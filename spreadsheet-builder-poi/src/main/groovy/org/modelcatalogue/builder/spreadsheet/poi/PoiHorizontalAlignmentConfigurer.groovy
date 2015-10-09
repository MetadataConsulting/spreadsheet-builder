package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.modelcatalogue.builder.spreadsheet.api.HorizontalAlignmentConfigurer

@CompileStatic class PoiHorizontalAlignmentConfigurer implements HorizontalAlignmentConfigurer {

    final PoiCellStyle style

    PoiHorizontalAlignmentConfigurer(PoiCellStyle style) {
        this.style = style
    }

    @Override
    Object getRight() {
        style.setHorizontalAlignment(HorizontalAlignment.RIGHT)
        return null
    }

    @Override
    Object getLeft() {
        style.setHorizontalAlignment(HorizontalAlignment.LEFT)
        return null
    }

    @Override
    Object getGeneral() {
        style.setHorizontalAlignment(HorizontalAlignment.GENERAL)
        return null
    }

    @Override
    Object getCenter() {
        style.setHorizontalAlignment(HorizontalAlignment.CENTER)
        return null
    }

    @Override
    Object getFill() {
        style.setHorizontalAlignment(HorizontalAlignment.FILL)
        return null
    }

    @Override
    Object getJustify() {
        style.setHorizontalAlignment(HorizontalAlignment.JUSTIFY)
        return null
    }

    @Override
    Object getCenterSelection() {
        style.setHorizontalAlignment(HorizontalAlignment.CENTER_SELECTION)
        return null
    }
}
