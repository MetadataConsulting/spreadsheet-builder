package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.ForegroundFill;
import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;
import org.modelcatalogue.spreadsheet.query.api.BorderCriterion;
import org.modelcatalogue.spreadsheet.query.api.CellStyleCriterion;
import org.modelcatalogue.spreadsheet.query.api.Predicate;
import org.modelcatalogue.spreadsheet.query.api.FontCriterion;

final class SimpleCellStyleCriterion implements CellStyleCriterion {

    private final SimpleCellCriterion parent;

    SimpleCellStyleCriterion(SimpleCellCriterion parent) {
        this.parent = parent;
    }

    @Override
    public void background(final String hexColor) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && new Color(hexColor).equals(o.getStyle().getBackground());
            }
        });
    }

    @Override
    public void background(final Color color) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && color.equals(o.getStyle().getBackground());
            }
        });
    }

    @Override
    public void background(final Predicate<Color> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getBackground());
            }
        });
    }

    @Override
    public void foreground(final String hexColor) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && new Color(hexColor).equals(o.getStyle().getForeground());
            }
        });
    }

    @Override
    public void foreground(final Color color) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && color.equals(o.getStyle().getForeground());
            }
        });
    }

    @Override
    public void foreground(final Predicate<Color> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getForeground());
            }
        });
    }

    @Override
    public void fill(final ForegroundFill fill) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && fill.equals(o.getStyle().getFill());
            }
        });
    }

    @Override
    public void fill(final Predicate<ForegroundFill> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getFill());
            }
        });
    }

    @Override
    public void indent(final int indent) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && indent == o.getStyle().getIndent();
            }
        });
    }

    @Override
    public void indent(final Predicate<Integer> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getIndent());
            }
        });
    }

    @Override
    public void rotation(final int rotation) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && rotation == o.getStyle().getRotation();
            }
        });
    }

    @Override
    public void rotation(final Predicate<Integer> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getRotation());
            }
        });
    }

    @Override
    public void format(final String format) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && format.equals(o.getStyle().getFormat());
            }
        });
    }

    @Override
    public void format(final Predicate<String> format) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && format.test(o.getStyle().getFormat());
            }
        });
    }

    @Override
    public void font(Configurer<FontCriterion> fontCriterion) {
        SimpleFontCriterion simpleFontCriterion = new SimpleFontCriterion(parent);
        fontCriterion.configure(simpleFontCriterion);
    }

    @Override
    public void border(Configurer<BorderCriterion> borderConfiguration) {
        border(Keywords.BorderSide.BORDER_SIDES, borderConfiguration);
    }

    @Override
    public void border(Keywords.BorderSide location, Configurer<BorderCriterion> borderConfiguration) {
        border(new Keywords.BorderSide[] {location}, borderConfiguration);
    }

    @Override
    public void border(Keywords.BorderSide first, Keywords.BorderSide second, Configurer<BorderCriterion> borderConfiguration) {
        border(new Keywords.BorderSide[] {first, second}, borderConfiguration);
    }

    @Override
    public void border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, Configurer<BorderCriterion> borderConfiguration) {
        border(new Keywords.BorderSide[] {first, second, third}, borderConfiguration);
    }

    private void border(Keywords.BorderSide[] sides, Configurer<BorderCriterion> borderConfiguration) {
        for (Keywords.BorderSide side : sides) {
            SimpleBorderCriterion criterion = new SimpleBorderCriterion(parent, side);
            borderConfiguration.configure(criterion);
        }
    }

    @Override
    public ForegroundFill getNoFill() {
        return ForegroundFill.NO_FILL;
    }

    @Override
    public ForegroundFill getSolidForeground() {
        return ForegroundFill.SOLID_FOREGROUND;
    }

    @Override
    public ForegroundFill getFineDots() {
        return ForegroundFill.FINE_DOTS;
    }

    @Override
    public ForegroundFill getAltBars() {
        return ForegroundFill.ALT_BARS;
    }

    @Override
    public ForegroundFill getSparseDots() {
        return ForegroundFill.SPARSE_DOTS;
    }

    @Override
    public ForegroundFill getThickHorizontalBands() {
        return ForegroundFill.THICK_HORZ_BANDS;
    }

    @Override
    public ForegroundFill getThickVerticalBands() {
        return ForegroundFill.THICK_VERT_BANDS;
    }

    @Override
    public ForegroundFill getThickBackwardDiagonals() {
        return ForegroundFill.THICK_BACKWARD_DIAG;
    }

    @Override
    public ForegroundFill getThickForwardDiagonals() {
        return ForegroundFill.THICK_FORWARD_DIAG;
    }

    @Override
    public ForegroundFill getBigSpots() {
        return ForegroundFill.BIG_SPOTS;
    }

    @Override
    public ForegroundFill getBricks() {
        return ForegroundFill.BRICKS;
    }

    @Override
    public ForegroundFill getThinHorizontalBands() {
        return ForegroundFill.THIN_HORZ_BANDS;
    }

    @Override
    public ForegroundFill getThinVerticalBands() {
        return ForegroundFill.THIN_VERT_BANDS;
    }

    @Override
    public ForegroundFill getThinBackwardDiagonals() {
        return ForegroundFill.THIN_BACKWARD_DIAG;
    }

    @Override
    public ForegroundFill getThinForwardDiagonals() {
        return ForegroundFill.THICK_FORWARD_DIAG;
    }

    @Override
    public ForegroundFill getSquares() {
        return ForegroundFill.SQUARES;
    }

    @Override
    public ForegroundFill getDiamonds() {
        return ForegroundFill.DIAMONDS;
    }

    @Override
    public Keywords.PureBorderSide getLeft() {
        return Keywords.PureBorderSide.LEFT;
    }

    @Override
    public Keywords.PureBorderSide getRight() {
        return Keywords.PureBorderSide.RIGHT;
    }

    @Override
    public Keywords.BorderSideAndVerticalAlignment getTop() {
        return Keywords.BorderSideAndVerticalAlignment.TOP;
    }

    @Override
    public Keywords.BorderSideAndVerticalAlignment getBottom() {
        return Keywords.BorderSideAndVerticalAlignment.BOTTOM;
    }

}
