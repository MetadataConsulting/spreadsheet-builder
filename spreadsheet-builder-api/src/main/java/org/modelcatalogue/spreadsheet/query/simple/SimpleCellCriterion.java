package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Comment;
import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.query.api.CellCriterion;
import org.modelcatalogue.spreadsheet.query.api.CellStyleCriterion;
import org.modelcatalogue.spreadsheet.query.api.Predicate;

import java.util.Calendar;
import java.util.Date;

final class SimpleCellCriterion extends AbstractCriterion<Cell, CellCriterion> implements CellCriterion {

    SimpleCellCriterion() {}

    private SimpleCellCriterion(boolean disjoint) {
        super(disjoint);
    }

    @Override
    public SimpleCellCriterion date(final Date value) {
        addValueCondition(value, Date.class);
        return this;
    }

    @Override
    public SimpleCellCriterion date(final Predicate<Date> predicate) {
        addValueCondition(predicate, Date.class);
        return this;
    }

    @Override
    public SimpleCellCriterion number(Double value) {
        addValueCondition(value, Double.class);
        return this;
    }

    @Override
    public SimpleCellCriterion number(Predicate<Double> predicate) {
        addValueCondition(predicate, Double.class);
        return this;
    }

    @Override
    public SimpleCellCriterion string(String value) {
        addValueCondition(value, String.class);
        return this;
    }

    @Override
    public SimpleCellCriterion string(Predicate<String> predicate) {
        addValueCondition(predicate, String.class);
        return this;
    }

    @Override
    public SimpleCellCriterion value(Object value) {
        if (value == null) {
            string("");
            return this;
        }
        if (value instanceof Date) {
            date((Date) value);
            return this;
        }
        if (value instanceof Calendar) {
            date(((Calendar) value).getTime());
            return this;
        }
        if (value instanceof Number) {
            number(((Number) value).doubleValue());
            return this;
        }
        if (value instanceof Boolean) {
            bool((Boolean) value);
        }
        string(value.toString());
        return this;
    }

    @Override
    public SimpleCellCriterion name(final String name) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return name.equals(o.getName());
            }
        });
        return this;
    }

    @Override
    public SimpleCellCriterion comment(final String comment) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return comment.equals(o.getComment().getText());
            }
        });
        return this;
    }

    @Override
    public SimpleCellCriterion bool(Boolean value) {
        addValueCondition(value, Boolean.class);
        return this;
    }


    @Override
    public SimpleCellCriterion style(Configurer<CellStyleCriterion> styleCriterion) {
        SimpleCellStyleCriterion criterion = new SimpleCellStyleCriterion(this);
        Configurer.Runner.doConfigure(styleCriterion, criterion);
        // no need to add criteria, they are added by the style criterion itself
        return this;
    }

    @Override
    public SimpleCellCriterion rowspan(final int span) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return span == o.getRowspan();
            }
        });
        return this;
    }

    @Override
    public SimpleCellCriterion rowspan(final Predicate<Integer> predicate) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return predicate.test(o.getRowspan());
            }
        });
        return this;
    }

    @Override
    public SimpleCellCriterion colspan(final int span) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return span == o.getColspan();
            }
        });
        return this;
    }

    @Override
    public SimpleCellCriterion colspan(final Predicate<Integer> predicate) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return predicate.test(o.getColspan());
            }
        });
        return this;
    }

    @Override
    public SimpleCellCriterion name(final Predicate<String> predicate) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return predicate.test(o.getName());
            }
        });
        return this;
    }

    @Override
    public SimpleCellCriterion comment(final Predicate<Comment> predicate) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return predicate.test(o.getComment());
            }
        });
        return this;
    }

    private <T> void addValueCondition(final T value, final Class<T> type) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                try {
                    return value.equals(o.read(type));
                } catch (Exception e) {
                    return false;
                }
            }
        });
    }

    private <T> void addValueCondition(final Predicate<T> predicate, final Class<T> type) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                try {
                    return predicate.test(o.read(type));
                } catch (Exception e) {
                    return false;
                }
            }
        });
    }

    @Override
    public SimpleCellCriterion or(Configurer<CellCriterion> sheetCriterion) {
        return (SimpleCellCriterion) super.or(sheetCriterion);
    }

    @Override
    public CellCriterion having(Predicate<Cell> cellPredicate) {
        addCondition(cellPredicate);
        return this;
    }

    @Override
    CellCriterion newDisjointCriterionInstance() {
        return new SimpleCellCriterion(true);
    }
}
