package org.modelcatalogue.spreadsheet.api;


public interface HorizontalAlignmentConfigurer {

    // following methods must return something otherwise they are not considered to be getters

    Object getRight();
    Object getLeft();

    Object getGeneral();
    Object getCenter();
    Object getFill();
    Object getJustify();
    Object getCenterSelection();

}
