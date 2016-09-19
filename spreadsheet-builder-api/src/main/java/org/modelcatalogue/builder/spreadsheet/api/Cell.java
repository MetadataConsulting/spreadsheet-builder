package org.modelcatalogue.builder.spreadsheet.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface Cell extends HasStyle {


    int getColumn();
    Row getRow();

    /**
     * Sets the value.
     * @param value new value
     */
    void value(Object value);
    void name(String name);
    void formula(String formula);
    void comment(String comment);
    void comment(@DelegatesTo(Comment.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Comment") Closure commentDefinition);

    LinkDefinition link(ToKeyword to);

    void colspan(int span);
    void rowspan(int span);

    /**
     * Sets the width as multiplier of standard character width.
     *
     * The width applies on the whole column.
     *
     * @param width the width as multiplier of standard character width
     */
    void width(double width);

    /**
     * Sets the height of the cell in points (multiples of 20 twips).
     *
     * The height applies on the whole row.
     *
     * @param height the height of the cell in points (multiples of 20 twips)
     */
    void height(double height);

    /**
     * Sets that the current column should have automatic width.
     * @param auto keyword
     */
    void width(AutoKeyword auto);

    AutoKeyword getAuto();
    ToKeyword getTo();

    /**
     * Add a new text run to the cell.
     *
     * This method can be called multiple times. The value of the cell will be result of appending all the text
     * values supplied.
     *
     * @param text new text run
     */
    void text(String text);

    /**
     * Add a new text run to the cell.
     *
     * This method can be called multiple times. The value of the cell will be result of appending all the text
     * values supplied.
     *
     * @param text new text run
     */
    void text(String text, @DelegatesTo(Font.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Font") Closure fontConfiguration);

    ImageCreator png(ImageKeyword image);
    ImageCreator jpeg(ImageKeyword image);
    ImageCreator pict(ImageKeyword image);
    ImageCreator emf(ImageKeyword image);
    ImageCreator wmf(ImageKeyword image);
    ImageCreator dib(ImageKeyword image);

    ImageKeyword getImage();

}
