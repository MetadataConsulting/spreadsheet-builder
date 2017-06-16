package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.query.api.BorderCriterion;
import org.modelcatalogue.spreadsheet.query.api.Predicate;

final class SimpleBorderCriterion implements BorderCriterion {

    private final SimpleCellCriterion parent;
    private final Keywords.BorderSide side;

    SimpleBorderCriterion(SimpleCellCriterion parent, Keywords.BorderSide side) {
        this.parent = parent;
        this.side = side;
    }

    @Override
    public SimpleBorderCriterion style(final BorderStyle borderStyle) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Border border = style.getBorder(side);
                return border != null && borderStyle.equals(border.getStyle());
            }
        });
        return this;
    }

    @Override
    public SimpleBorderCriterion style(final Predicate<BorderStyle> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Border border = style.getBorder(side);
                return border != null && predicate.test(border.getStyle());
            }
        });
        return this;
    }

    @Override
    public SimpleBorderCriterion color(String hexColor) {
        color(new Color(hexColor));
        return this;
    }

    @Override
    public SimpleBorderCriterion color(final Color color) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Border border = style.getBorder(side);
                return border != null && color.equals(border.getColor());
            }
        });
        return this;
    }

    @Override
    public SimpleBorderCriterion color(final Predicate<Color> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Border border = style.getBorder(side);
                return border != null && predicate.test(border.getColor());
            }
        });
        return this;
    }
}
