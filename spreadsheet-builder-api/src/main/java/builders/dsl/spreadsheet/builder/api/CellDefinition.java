package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Keywords;

import java.util.function.Consumer;

public interface CellDefinition extends HasStyle {

    /**
     * Sets the value.
     * @param value new value
     */
    CellDefinition value(Object value);
    CellDefinition name(String name);
    CellDefinition formula(String formula);
    CellDefinition comment(String comment);
    CellDefinition comment(Consumer<CommentDefinition> commentDefinition);

    LinkDefinition link(Keywords.To to);

    CellDefinition colspan(int span);
    CellDefinition rowspan(int span);

    /**
     * Sets the width as multiplier of standard character width.
     *
     * The width applies on the whole column.
     *
     * @param width the width as multiplier of standard character width
     * @return dimension modifier which allows to recalculate the number set to cm or inches
     */
    DimensionModifier width(double width);

    /**
     * Sets the height of the cell in points (multiples of 20 twips).
     *
     * The height applies on the whole row.
     *
     * @param height the height of the cell in points (multiples of 20 twips)
     * @return dimension modifier which allows to recalculate the number set to cm or inches
     */
    DimensionModifier height(double height);

    /**
     * Sets that the current column should have automatic width.
     * @param auto keyword
     */
    CellDefinition width(Keywords.Auto auto);

    /**
     * Add a new text run to the cell.
     *
     * This method can be called multiple times. The value of the cell will be result of appending all the text
     * values supplied.
     *
     * @param text new text run
     */
    CellDefinition text(String text);

    /**
     * Add a new text run to the cell.
     *
     * This method can be called multiple times. The value of the cell will be result of appending all the text
     * values supplied.
     *
     * @param text new text run
     */
    CellDefinition text(String text, Consumer<FontDefinition> fontConfiguration);

    ImageCreator png(Keywords.Image image);
    ImageCreator jpeg(Keywords.Image image);
    ImageCreator pict(Keywords.Image image);
    ImageCreator emf(Keywords.Image image);
    ImageCreator wmf(Keywords.Image image);
    ImageCreator dib(Keywords.Image image);

    CellDefinition style(String name, Consumer<CellStyleDefinition> styleDefinition);
    CellDefinition styles(Iterable<String> names, Consumer<CellStyleDefinition> styleDefinition);
    CellDefinition styles(Iterable<String> styles, Iterable<Consumer<CellStyleDefinition>> styleDefinitions);
    CellDefinition style(Consumer<CellStyleDefinition> styleDefinition);
    CellDefinition style(String name);
    CellDefinition styles(String... names);
    CellDefinition styles(Iterable<String> names);

}
