package builders.dsl.spreadsheet.impl;

import builders.dsl.spreadsheet.builder.api.CellDefinition;
import builders.dsl.spreadsheet.builder.api.Resolvable;

/**
 * Pending link is a link which needs to be resolved at the end of the build when all the named references are known.
 */
public abstract class AbstractPendingLink implements Resolvable {

    protected AbstractPendingLink(CellDefinition cell, final String name) {
        if (!name.equals(Utils.fixName(name))) {
            throw new IllegalArgumentException("Cannot call cell \'" + name + "\' as this is invalid identifier in Excel. Suggestion: " + Utils.fixName(name));
        }

        this.cell = cell;
        this.name = name;
    }

    public final CellDefinition getCell() {
        return cell;
    }

    public final String getName() {
        return name;
    }

    private final CellDefinition cell;
    private final String name;
}
