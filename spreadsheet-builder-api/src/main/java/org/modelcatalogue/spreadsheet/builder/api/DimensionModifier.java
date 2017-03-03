package org.modelcatalogue.spreadsheet.builder.api;

/**
 * Allows to alter the height of the row in centimeters or
 */
public interface DimensionModifier {

    /**
     * Converts the dimension to centimeters.
     *
     * This feature is currently experimental.
     * @return null to comply with getter signatures
     */
    Object getCm();

    /**
     * Converts the dimension to inches.
     *
     * This feature is currently experimental.
     * @return null to comply with getter signatures
     */
    Object getInch();

    /**
     * Converts the dimension to inches.
     *
     * This feature is currently experimental.
     * @return null to comply with getter signatures
     */
    Object getInches();
}
