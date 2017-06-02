package org.modelcatalogue.spreadsheet.builder.api;

/**
 * Allows to alter the height of the row in centimeters or
 */
public interface DimensionModifier {

    /**
     * Converts the dimension to centimeters.
     *
     * This feature is currently experimental.
     */
    CellDefinition getCm();

    /**
     * Converts the dimension to inches.
     *
     * This feature is currently experimental.
     */
    CellDefinition getInch();

    /**
     * Converts the dimension to inches.
     *
     * This feature is currently experimental.
     */
    CellDefinition getInches();

    /**
     * Keeps the dimesion in points.
     *
     * This feature is currently experimental.
     */
    CellDefinition getPoints();
}
