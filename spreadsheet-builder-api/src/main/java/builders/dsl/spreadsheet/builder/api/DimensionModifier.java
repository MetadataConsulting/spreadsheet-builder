package builders.dsl.spreadsheet.builder.api;

/**
 * Allows to alter the height of the row in centimeters or
 */
public interface DimensionModifier {

    /**
     * Converts the dimension to centimeters.
     *
     * This feature is currently experimental.
     */
    CellDefinition cm();

    /**
     * Converts the dimension to inches.
     *
     * This feature is currently experimental.
     */
    CellDefinition inch();

    /**
     * Converts the dimension to inches.
     *
     * This feature is currently experimental.
     */
    CellDefinition inches();

    /**
     * Keeps the dimesion in points.
     *
     * This feature is currently experimental.
     */
    CellDefinition points();
}
