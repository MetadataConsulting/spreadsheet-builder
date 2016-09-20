package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.query.api.Condition;
import org.modelcatalogue.spreadsheet.query.api.FontCriterion;

import java.util.EnumSet;

class SimpleFontCriterion implements FontCriterion {

    private final SimpleCellCriterion parent;

    SimpleFontCriterion(SimpleCellCriterion parent) {
        this.parent = parent;
    }

    @Override
    public void color(String hexColor) {
        color(new Color(hexColor));
    }

    @Override
    public void color(final Color color) {
        parent.addCondition(new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && color.equals(font.getColor());
            }
        });
    }

    @Override
    public void color(final Condition<Color> conition) {
        parent.addCondition(new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && conition.evaluate(font.getColor());
            }
        });
    }

    @Override
    public void size(final int size) {
        parent.addCondition(new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && size == font.getSize();
            }
        });
    }

    @Override
    public void size(final Condition<Integer> condition) {
        parent.addCondition(new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && condition.evaluate(font.getSize());
            }
        });
    }

    @Override
    public void name(final String name) {
        parent.addCondition(new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && name.equals(font.getName());
            }
        });
    }

    @Override
    public void name(final Condition<String> condition) {
        parent.addCondition(new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && condition.evaluate(font.getName());
            }
        });
    }

    @Override
    public void make(final FontStyle first, final FontStyle... other) {
        parent.addCondition(new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                if (font == null) {
                    return false;
                }

                EnumSet<FontStyle> wanted = EnumSet.of(first, other);
                EnumSet<FontStyle> actual = font.getStyles();

                for (FontStyle fs : wanted) {
                    if (!actual.contains(fs)) {
                        return false;
                    }
                }

                return true;
            }
        });
    }

    @Override
    public void make(final Condition<EnumSet<FontStyle>> condition) {
        parent.addCondition(new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                CellStyle style = o.getStyle();
                if (style == null) {
                    return false;
                }
                Font font = style.getFont();
                return font != null && condition.evaluate(font.getStyles());
            }
        });
    }

    @Override
    public FontStyle getItalic() {
        return FontStyle.ITALIC;
    }

    @Override
    public FontStyle getBold() {
        return FontStyle.BOLD;
    }

    @Override
    public FontStyle getStrikeout() {
        return FontStyle.STRIKEOUT;
    }

    @Override
    public FontStyle getUnderline() {
        return FontStyle.UNDERLINE;
    }
}
