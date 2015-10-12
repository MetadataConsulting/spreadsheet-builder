package org.modelcatalogue.builder.spreadsheet.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface Cell extends HasStyle {

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


}
