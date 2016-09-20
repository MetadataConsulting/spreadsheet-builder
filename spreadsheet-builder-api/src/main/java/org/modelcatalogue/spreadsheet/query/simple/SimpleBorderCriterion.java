package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.query.api.BorderCriterion;
import org.modelcatalogue.spreadsheet.query.api.Predicate;

class SimpleBorderCriterion extends AbstractBorderProvider implements BorderCriterion {

    private final SimpleCellCriterion parent;
    private final Keywords.BorderSide side;

    SimpleBorderCriterion(SimpleCellCriterion parent, Keywords.BorderSide side) {
        this.parent = parent;
        this.side = side;
    }

    @Override
    public void style(final BorderStyle borderStyle) {
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
    }

    @Override
    public void style(final Predicate<BorderStyle> predicate) {
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
    }

    @Override
    public void color(String hexColor) {
        color(new Color(hexColor));
    }

    @Override
    public void color(final Color color) {
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
    }

    @Override
    public void color(final Predicate<Color> predicate) {
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
    }
}
