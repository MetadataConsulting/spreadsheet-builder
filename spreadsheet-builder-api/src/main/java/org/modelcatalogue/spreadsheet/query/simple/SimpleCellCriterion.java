package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Comment;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;
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
    public void date(final Date value) {
        addValueCondition(value, Date.class);
    }

    @Override
    public void date(final Predicate<Date> predicate) {
        addValueCondition(predicate, Date.class);
    }

    @Override
    public void number(Double value) {
        addValueCondition(value, Double.class);
    }

    @Override
    public void number(Predicate<Double> predicate) {
        addValueCondition(predicate, Double.class);
    }

    @Override
    public void string(String value) {
        addValueCondition(value, String.class);
    }

    @Override
    public void string(Predicate<String> predicate) {
        addValueCondition(predicate, String.class);
    }

    @Override
    public void value(Object value) {
        if (value == null) {
            string("");
            return;
        }
        if (value instanceof Date) {
            date((Date) value);
            return;
        }
        if (value instanceof Calendar) {
            date(((Calendar) value).getTime());
            return;
        }
        if (value instanceof Number) {
            number(((Number) value).doubleValue());
            return;
        }
        if (value instanceof Boolean) {
            bool((Boolean) value);
        }
        string(value.toString());
    }

    @Override
    public void name(final String name) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return name.equals(o.getName());
            }
        });
    }

    @Override
    public void comment(final String comment) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return comment.equals(o.getComment().getText());
            }
        });
    }

    @Override
    public void bool(Boolean value) {
        addValueCondition(value, Boolean.class);
    }


    @Override
    public void style(Configurer<CellStyleCriterion> styleCriterion) {
        SimpleCellStyleCriterion criterion = new SimpleCellStyleCriterion(this);
        Configurer.Runner.doConfigure(styleCriterion, criterion);
        // no need to add criteria, they are added by the style criterion itself
    }

    @Override
    public void rowspan(final int span) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return span == o.getRowspan();
            }
        });
    }

    @Override
    public void rowspan(final Predicate<Integer> predicate) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return predicate.test(o.getRowspan());
            }
        });
    }

    @Override
    public void colspan(final int span) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return span == o.getColspan();
            }
        });
    }

    @Override
    public void colspan(final Predicate<Integer> predicate) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return predicate.test(o.getColspan());
            }
        });
    }

    @Override
    public void name(final Predicate<String> predicate) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return predicate.test(o.getName());
            }
        });
    }

    @Override
    public void comment(final Predicate<Comment> predicate) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return predicate.test(o.getComment());
            }
        });
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
    CellCriterion newDisjointCriterionInstance() {
        return new SimpleCellCriterion(true);
    }
}
