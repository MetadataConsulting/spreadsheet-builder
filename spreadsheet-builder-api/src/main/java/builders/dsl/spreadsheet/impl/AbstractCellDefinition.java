package builders.dsl.spreadsheet.impl;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.api.Spannable;
import builders.dsl.spreadsheet.builder.api.*;

import java.util.*;
import java.util.function.Consumer;

public abstract class AbstractCellDefinition implements CellDefinition, Resolvable, Spannable {
    
    protected AbstractCellDefinition(AbstractRowDefinition row) {
        this.row = checkNotNull(row, "Row");
    }

    private static <T> T checkNotNull(T o, String what) {
        if (o == null) {
            throw new IllegalArgumentException(what + " cannot be null");
        }

        return o;
    }

    @Override
    public final CellDefinition comment(final String commentText) {
        comment(commentDefinition -> commentDefinition.text(commentText));
        return this;
    }

    @Override
    public final CellDefinition formula(String formula) {
        row.getSheet().getWorkbook().addPendingFormula(createPendingFormula(formula));
        return this;
    }

    protected abstract AbstractPendingFormula createPendingFormula(String formula);

    @Override
    public final CellDefinition comment(Consumer<CommentDefinition> commentDefinition) {
        DefaultCommentDefinition poiComment = new DefaultCommentDefinition();
        commentDefinition.accept(poiComment);
        applyComment(poiComment);
        return this;
    }

    @Override
    public final CellDefinition colspan(int span) {
        this.colspan = span;
        return this;
    }

    @Override
    public final CellDefinition rowspan(int span) {
        this.rowspan = span;
        return this;
    }

    @Override
    public final CellDefinition style(String name) {
        return styles(Collections.singleton(name), Collections.<Consumer<CellStyleDefinition>>emptyList());
    }

    @Override
    public final CellDefinition styles(String... names) {
        return styles(Arrays.asList(names), Collections.<Consumer<CellStyleDefinition>>emptyList());
    }

    @Override
    public final CellDefinition style(Consumer<CellStyleDefinition> styleDefinition) {
        return styles(Collections.<String>emptyList(), Collections.singleton(styleDefinition));
    }

    @Override
    public final CellDefinition styles(Iterable<String> names) {
        return styles(names, Collections.<Consumer<CellStyleDefinition>>emptyList());
    }

    @Override
    public final CellDefinition style(String name, Consumer<CellStyleDefinition> styleDefinition) {
        return styles(Collections.singleton(name), Collections.singleton(styleDefinition));
    }

    @Override
    public final CellDefinition styles(Iterable<String> names, Consumer<CellStyleDefinition> styleDefinition) {
        return styles(names, Collections.singleton(styleDefinition));
    }

    @Override
    public final CellDefinition styles(Iterable<String> names, Iterable<Consumer<CellStyleDefinition>> styleDefinition) {
        if (styleDefinition == null || !styleDefinition.iterator().hasNext()) {
            if (names == null || !names.iterator().hasNext()) {
                return this;
            }

            Set<String> allNames = new LinkedHashSet<String>();
            for (String name: names) {
                allNames.add(name);
            }
            allNames.addAll(row.getStyles());

            if (cellStyle == null) {
                cellStyle = row.getSheet().getWorkbook().getStyles(allNames);
                assignStyle(cellStyle);
                return this;
            }

            if (cellStyle.isSealed()) {
                if (!row.getStyles().isEmpty()) {
                    cellStyle = null;
                    styles(allNames);
                    return this;
                }
            } else {
                for (String name : names) {
                    row.getSheet().getWorkbook().getStyleDefinition(name).accept(cellStyle);
                }
            }
            return this;
        }

        if (cellStyle == null) {
            cellStyle = createCellStyle();
        }

        if (cellStyle.isSealed()) {
            throw new IllegalStateException("The cell style '" + Utils.join(names, ".") + "' is already sealed! You need to create new style. Use 'styles' method to combine multiple named styles! Create new named style if you're trying to update existing style with closure definition.");
        }

        for (String name : names) {
            row.getSheet().getWorkbook().getStyleDefinition(name).accept(cellStyle);
        }

        for (Consumer<CellStyleDefinition> configurer : styleDefinition) {
            configurer.accept(cellStyle);
        }

        return this;
    }

    protected abstract AbstractCellStyleDefinition createCellStyle();

    protected abstract void assignStyle(CellStyleDefinition cellStyle);

    @Override
    public final CellDefinition name(final String name) {
        if (!Utils.fixName(name).equals(name)) {
            throw new IllegalArgumentException("Name " + name + " is not valid Excel name! Suggestion: " + Utils.fixName(name));
        }

        doName(name);
        return this;
    }

    protected abstract void doName(String name);

    @Override
    public final LinkDefinition link(Keywords.To to) {
        return createLinkDefinition();
    }

    protected abstract LinkDefinition createLinkDefinition();

    @Override
    public final CellDefinition text(String run) {
        text(run, null);
        return this;
    }

    @Override
    public final CellDefinition text(String run, Consumer<FontDefinition> fontConfiguration) {
        if (run == null || run.length() == 0) {
            return this;
        }

        int start = 0;
        if (richTextParts != null && richTextParts.size() > 0) {
            start = richTextParts.get(richTextParts.size() - 1).getEnd();
        }

        int end = start + run.length();

        if (fontConfiguration == null) {
            richTextParts.add(new RichTextPart(run, null, start, end));
            return this;
        }


        FontDefinition font = createFontDefinition();
        fontConfiguration.accept(font);

        richTextParts.add(new RichTextPart(run, font, start, end));
        return this;
    }

    protected abstract FontDefinition createFontDefinition();

    public final int getColspan() {
        return colspan;
    }

    public final int getRowspan() {
        return rowspan;
    }

    protected abstract void applyComment(DefaultCommentDefinition comment);

    public AbstractRowDefinition getRow() {
        return row;
    }

    private final AbstractRowDefinition row;
    private int colspan = 1;
    private int rowspan = 1;
    protected AbstractCellStyleDefinition cellStyle;
    protected List<RichTextPart> richTextParts = new ArrayList<RichTextPart>();
}
