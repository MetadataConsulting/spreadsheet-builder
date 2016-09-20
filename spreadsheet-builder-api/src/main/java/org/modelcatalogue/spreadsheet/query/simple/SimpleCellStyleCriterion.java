package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.ForegroundFill;
import org.modelcatalogue.spreadsheet.query.api.CellStyleCriterion;
import org.modelcatalogue.spreadsheet.query.api.Predicate;
import org.modelcatalogue.spreadsheet.query.api.FontCriterion;

class SimpleCellStyleCriterion implements CellStyleCriterion {

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
    public void font(@DelegatesTo(FontCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.FontCriterion") Closure fontCriterion) {
        SimpleFontCriterion simpleFontCriterion = new SimpleFontCriterion(parent);
        DefaultGroovyMethods.with(simpleFontCriterion, fontCriterion);
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

}
