package builders.dsl.spreadsheet.impl;

import builders.dsl.spreadsheet.builder.api.CellDefinition;
import builders.dsl.spreadsheet.builder.api.Resolvable;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Pending formula is a formula definition which needs to be resolved at the end of the build where all named references
 * are know.
 */
public abstract class AbstractPendingFormula implements Resolvable {

    protected AbstractPendingFormula(CellDefinition cell, String formula) {
        this.cell = cell;
        this.formula = formula;
    }

    public final void resolve() {
        String expandedFormula = expandNames(this.formula);
        doResolve(expandedFormula);
    }

    protected abstract void doResolve(String expandedFormula);

    private String expandNames(String withNames) {
        final Matcher matcher = Pattern.compile("#\\{(.+?)\\}").matcher(withNames);
        if (matcher.find()) {
            final StringBuffer sb = new StringBuffer(withNames.length() + 16);
            do {
                String replacement = findRefersToFormula(matcher.group(1));
                matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
            } while (matcher.find());
            matcher.appendTail(sb);
            return sb.toString();
        }
        return withNames;
    }

    protected abstract String findRefersToFormula(final String name);

    public final CellDefinition getCell() {
        return cell;
    }

    public final String getFormula() {
        return formula;
    }

    private final CellDefinition cell;
    private final String formula;
}
