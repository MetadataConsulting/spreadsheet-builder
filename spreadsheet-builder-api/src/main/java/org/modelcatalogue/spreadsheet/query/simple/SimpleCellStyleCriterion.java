package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.ForegroundFill;
import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.api.Configurer;
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
    public SimpleCellStyleCriterion background(final String hexColor) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && new Color(hexColor).equals(o.getStyle().getBackground());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion background(final Color color) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && color.equals(o.getStyle().getBackground());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion background(final Predicate<Color> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getBackground());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion foreground(final String hexColor) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && new Color(hexColor).equals(o.getStyle().getForeground());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion foreground(final Color color) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && color.equals(o.getStyle().getForeground());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion foreground(final Predicate<Color> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getForeground());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion fill(final ForegroundFill fill) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && fill.equals(o.getStyle().getFill());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion fill(final Predicate<ForegroundFill> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getFill());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion indent(final int indent) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && indent == o.getStyle().getIndent();
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion indent(final Predicate<Integer> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getIndent());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion rotation(final int rotation) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && rotation == o.getStyle().getRotation();
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion rotation(final Predicate<Integer> predicate) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && predicate.test(o.getStyle().getRotation());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion format(final String format) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && format.equals(o.getStyle().getFormat());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion format(final Predicate<String> format) {
        parent.addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getStyle() != null && format.test(o.getStyle().getFormat());
            }
        });
        return this;
    }

    @Override
    public SimpleCellStyleCriterion font(Configurer<FontCriterion> fontCriterion) {
        SimpleFontCriterion simpleFontCriterion = new SimpleFontCriterion(parent);
        Configurer.Runner.doConfigure(fontCriterion, simpleFontCriterion);
        return this;
    }

    @Override
    public SimpleCellStyleCriterion border(Configurer<BorderCriterion> borderConfiguration) {
        border(Keywords.BorderSide.BORDER_SIDES, borderConfiguration);
        return this;
    }

    @Override
    public SimpleCellStyleCriterion border(Keywords.BorderSide location, Configurer<BorderCriterion> borderConfiguration) {
        border(new Keywords.BorderSide[] {location}, borderConfiguration);
        return this;
    }

    @Override
    public SimpleCellStyleCriterion border(Keywords.BorderSide first, Keywords.BorderSide second, Configurer<BorderCriterion> borderConfiguration) {
        border(new Keywords.BorderSide[] {first, second}, borderConfiguration);
        return this;
    }

    @Override
    public SimpleCellStyleCriterion border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, Configurer<BorderCriterion> borderConfiguration) {
        border(new Keywords.BorderSide[] {first, second, third}, borderConfiguration);
        return this;
    }

    private void border(Keywords.BorderSide[] sides, Configurer<BorderCriterion> borderConfiguration) {
        for (Keywords.BorderSide side : sides) {
            SimpleBorderCriterion criterion = new SimpleBorderCriterion(parent, side);
            Configurer.Runner.doConfigure(borderConfiguration, criterion);
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
