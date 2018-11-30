package builders.dsl.spreadsheet.query.simple;

import builders.dsl.spreadsheet.api.*;
import builders.dsl.spreadsheet.query.api.BorderCriterion;
import java.util.function.Predicate;

final class SimpleBorderCriterion implements BorderCriterion {

    private final SimpleCellCriterion parent;
    private final Keywords.BorderSide side;

    SimpleBorderCriterion(SimpleCellCriterion parent, Keywords.BorderSide side) {
        this.parent = parent;
        this.side = side;
    }

    @Override
    public SimpleBorderCriterion style(final BorderStyle borderStyle) {
        parent.addCondition(o -> {
            CellStyle style = o.getStyle();
            if (style == null) {
                return false;
            }
            Border border = style.getBorder(side);
            return border != null && borderStyle.equals(border.getStyle());
        });
        return this;
    }

    @Override
    public SimpleBorderCriterion style(final Predicate<BorderStyle> predicate) {
        parent.addCondition(o -> {
            CellStyle style = o.getStyle();
            if (style == null) {
                return false;
            }
            Border border = style.getBorder(side);
            return border != null && predicate.test(border.getStyle());
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
        parent.addCondition(o -> {
            CellStyle style = o.getStyle();
            if (style == null) {
                return false;
            }
            Border border = style.getBorder(side);
            return border != null && color.equals(border.getColor());
        });
        return this;
    }

    @Override
    public SimpleBorderCriterion color(final Predicate<Color> predicate) {
        parent.addCondition(o -> {
            CellStyle style = o.getStyle();
            if (style == null) {
                return false;
            }
            Border border = style.getBorder(side);
            return border != null && predicate.test(border.getColor());
        });
        return this;
    }

    @Override
    public BorderCriterion having(final Predicate<Border> borderPredicate) {
        parent.addCondition(o -> {
            CellStyle style = o.getStyle();
            if (style == null) {
                return false;
            }
            Border border = style.getBorder(side);
            return border != null && borderPredicate.test(border);
        });
        return this;
    }
}
