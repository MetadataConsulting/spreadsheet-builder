package builders.dsl.spreadsheet.api.groovy;


import builders.dsl.spreadsheet.builder.api.CellStyleDefinition;

@Deprecated
public interface HorizontalAlignmentConfigurer {

    // following methods must return something otherwise they are not considered to be getters

    CellStyleDefinition getRight();
    CellStyleDefinition getLeft();

    CellStyleDefinition getGeneral();
    CellStyleDefinition getCenter();
    CellStyleDefinition getFill();
    CellStyleDefinition getJustify();
    CellStyleDefinition getCenterSelection();

}
