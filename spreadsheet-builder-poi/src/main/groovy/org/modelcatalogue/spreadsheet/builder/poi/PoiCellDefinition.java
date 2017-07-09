package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.builder.api.*;
import org.modelcatalogue.spreadsheet.impl.DefaultCommentDefinition;
import org.modelcatalogue.spreadsheet.impl.HeightModifier;
import org.modelcatalogue.spreadsheet.impl.Utils;
import org.modelcatalogue.spreadsheet.impl.WidthModifier;

import java.util.*;

class PoiCellDefinition implements CellDefinition, Resolvable, Cell {
    PoiCellDefinition(PoiRowDefinition row, XSSFCell xssfCell) {
        this.xssfCell = checkNotNull(xssfCell, "Cell");
        this.row = checkNotNull(row, "Row");
    }

    private static <T> T checkNotNull(T o, String what) {
        if (o == null) {
            throw new IllegalArgumentException(what + " cannot be null");
        }

        return o;
    }

    @Override
    public int getColumn() {
        return xssfCell.getColumnIndex() + 1;
    }

    @Override
    public String getColumnAsString() {
        return Util.toColumn(getColumn());
    }

    @Override
    public <T> T read(Class<T> type) {
        if (CharSequence.class.isAssignableFrom(type)) {
            return type.cast(xssfCell.getStringCellValue());
        }


        if (Date.class.isAssignableFrom(type)) {
            return type.cast(xssfCell.getDateCellValue());
        }


        if (Boolean.class.isAssignableFrom(type)) {
            return type.cast(xssfCell.getBooleanCellValue());
        }


        if (Number.class.isAssignableFrom(type)) {
            Double val = xssfCell.getNumericCellValue();
            return type.cast(val);
        }

        if (xssfCell.getRawValue() == null) {
            return null;
        }

        throw new IllegalArgumentException("Cannot read value " + xssfCell.getRawValue() + " of cell as " + String.valueOf(type));
    }

    @Override
    public Object getValue() {
        switch (xssfCell.getCellType()) {
            case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK:
                return "";
            case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BOOLEAN:
                return xssfCell.getBooleanCellValue();
            case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_ERROR:
                return xssfCell.getErrorCellString();
            case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA:
                return xssfCell.getCellFormula();
            case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC:
                return xssfCell.getNumericCellValue();
            case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING:
                return xssfCell.getStringCellValue();
        }
        return xssfCell.getRawValue();
    }

    @Override
    public PoiCellDefinition value(Object value) {
        if (value == null) {
            xssfCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK);
            return this;
        }

        if (value instanceof Number) {
            xssfCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC);
            xssfCell.setCellValue(((Number) value).doubleValue());
            return this;
        }

        if (value instanceof Date) {
            xssfCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC);
            xssfCell.setCellValue((Date) value);
            return this;
        }

        if (value instanceof Calendar) {
            xssfCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC);
            xssfCell.setCellValue((Calendar) value);
            return this;
        }

        if (value instanceof Boolean) {
            xssfCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BOOLEAN);
            xssfCell.setCellValue((Boolean) value);
            return this;
        }

        xssfCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING);
        xssfCell.setCellValue(value.toString());
        return this;
    }

    @Override
    public PoiCellDefinition style(Configurer<CellStyleDefinition> styleDefinition) {
        if (poiCellStyle == null) {
            poiCellStyle = new PoiCellStyleDefinition(this);
        }

        poiCellStyle.checkSealed();
        Configurer.Runner.doConfigure(styleDefinition, poiCellStyle);
        return this;
    }

    @Override
    public PoiCellDefinition comment(final String commentText) {
        comment(new Configurer<CommentDefinition>() {
            @Override
            public void configure(CommentDefinition commentDefinition) {
                commentDefinition.text(commentText);
            }
        });
        return this;
    }

    @Override
    public PoiCellDefinition formula(String formula) {
        row.getSheet().getWorkbook().addPendingFormula(formula, this);
        return this;
    }

    @Override
    public PoiCellDefinition comment(Configurer<CommentDefinition> commentDefinition) {
        DefaultCommentDefinition poiComment = new DefaultCommentDefinition();
        Configurer.Runner.doConfigure(commentDefinition, poiComment);
        applyTo(poiComment, xssfCell);
        return this;
    }

    @Override
    public Comment getComment() {
        XSSFComment comment = xssfCell.getCellComment();
        if (comment == null) {
            return new DefaultCommentDefinition();
        }

        DefaultCommentDefinition definition = new DefaultCommentDefinition();
        definition.author(comment.getAuthor());
        definition.text(comment.getString().getString());
        return definition;
    }

    @Override
    public CellStyle getStyle() {
        XSSFCellStyle cellStyle = xssfCell.getCellStyle();
        return cellStyle != null ? new PoiCellStyle(cellStyle) : null;
    }

    @Override
    public PoiCellDefinition colspan(int span) {
        this.colspan = span;
        return this;
    }

    @Override
    public PoiCellDefinition rowspan(int span) {
        this.rowspan = span;
        return this;
    }

    @Override
    public PoiCellDefinition style(String name) {
        styles(name);
        return this;
    }

    @Override
    public PoiCellDefinition styles(String... names) {
        styles(Arrays.asList(names));
        return this;
    }

    @Override
    public PoiCellDefinition styles(Iterable<String> names) {
        Set<String> allNames = new LinkedHashSet<String>();
        for (String name: names) {
            allNames.add(name);
        }
        allNames.addAll(row.getStyles());

        if (poiCellStyle == null) {
            poiCellStyle = (PoiCellStyleDefinition) row.getSheet().getWorkbook().getStyles(allNames);
            poiCellStyle.assignTo(this);
            return this;
        }

        if (poiCellStyle.isSealed() && !row.getStyles().isEmpty()) {
            poiCellStyle = null;
            styles(allNames);
            return this;
        }

        poiCellStyle.checkSealed();
        for (String name : names) {
            Configurer.Runner.doConfigure(row.getSheet().getWorkbook().getStyleDefinition(name), poiCellStyle);
        }

        return this;
    }

    @Override
    public PoiCellDefinition style(String name, Configurer<CellStyleDefinition> styleDefinition) {
        style(row.getSheet().getWorkbook().getStyleDefinition(name));
        style(styleDefinition);
        return this;
    }

    @Override
    public PoiCellDefinition styles(Iterable<String> names, Configurer<CellStyleDefinition> styleDefinition) {
        for (String name : names) {
            Configurer.Runner.doConfigure(row.getSheet().getWorkbook().getStyleDefinition(name), poiCellStyle);
        }

        style(styleDefinition);
        return this;
    }

    @Override
    public PoiCellDefinition styles(Iterable<String> names, Iterable<Configurer<CellStyleDefinition>> styleDefinition) {
        if (styleDefinition == null || !styleDefinition.iterator().hasNext()) {
            if (names == null || !names.iterator().hasNext()) {
                return this;
            }

            Set<String> allNames = new LinkedHashSet<String>();
            for (String name: names) {
                allNames.add(name);
            }
            allNames.addAll(row.getStyles());

            if (poiCellStyle == null) {
                poiCellStyle = (PoiCellStyleDefinition) row.getSheet().getWorkbook().getStyles(allNames);
                poiCellStyle.assignTo(this);
                return this;
            }

            if (poiCellStyle.isSealed() && !row.getStyles().isEmpty()) {
                poiCellStyle = null;
                styles(allNames);
                return this;
            }

        }

        if (poiCellStyle == null) {
            poiCellStyle = new PoiCellStyleDefinition(this);
        }

        poiCellStyle.checkSealed();
        for (String name : names) {
            Configurer.Runner.doConfigure(row.getSheet().getWorkbook().getStyleDefinition(name), poiCellStyle);
        }

        if (styleDefinition != null) {
            for (Configurer<CellStyleDefinition> configurer : styleDefinition) {
                Configurer.Runner.doConfigure(configurer, poiCellStyle);
            }
        }

        return this;
    }

    @Override
    public PoiCellDefinition name(final String name) {
        if (!Utils.fixName(name).equals(name)) {
            throw new IllegalArgumentException("Name " + name + " is not valid Excel name! Suggestion: " + Utils.fixName(name));
        }

        XSSFName theName = xssfCell.getRow().getSheet().getWorkbook().createName();
        theName.setNameName(name);
        theName.setRefersToFormula(generateRefersToFormula());
        return this;
    }

    private String generateRefersToFormula() {
        return "\'" + xssfCell.getSheet().getSheetName().replaceAll("'", "\\'") + "\'!" + xssfCell.getReference();
    }

    @Override
    public String getName() {
        Workbook wb = xssfCell.getSheet().getWorkbook();

        new CellReference(xssfCell).formatAsString();
        List<String> possibleReferences = new ArrayList<String>(Arrays.asList(new CellReference(xssfCell).formatAsString(), generateRefersToFormula()));
        for (int nn = 0; nn < wb.getNumberOfNames(); nn++) {
            Name n = ((XSSFWorkbook) wb).getNameAt(nn);
            for (String reference : possibleReferences) {
                if (n.getSheetIndex() == -1 || n.getSheetIndex() == wb.getSheetIndex(xssfCell.getSheet())) {
                    if (n.getRefersToFormula().equals(reference)) {
                        return n.getNameName();
                    }

                }

            }
        }

        return null;
    }

    @Override
    public LinkDefinition link(Keywords.To to) {
        return new PoiLinkDefinition(row.getSheet().getWorkbook(), this);
    }

    @Override
    public DimensionModifier width(double width) {
        row.getSheet().getSheet().setColumnWidth(xssfCell.getColumnIndex(), (int) Math.round(width * 255D));
        return new WidthModifier(this, width, WIDTH_POINTS_PER_CM, WIDTH_POINTS_PER_INCH);
    }

    @Override
    public DimensionModifier height(double height) {
        row.getRow().setHeightInPoints((float) height);
        return new HeightModifier(this, height, HEIGHT_POINTS_PER_CM, HEIGHT_POINTS_PER_INCH);
    }

    @Override
    public PoiCellDefinition width(Keywords.Auto auto) {
        row.getSheet().addAutoColumn(xssfCell.getColumnIndex());
        return this;
    }

    @Override
    public PoiCellDefinition text(String run) {
        text(run, null);
        return this;
    }

    @Override
    public PoiCellDefinition text(String run, Configurer<FontDefinition> fontConfiguration) {
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


        PoiFontDefinition font = new PoiFontDefinition(xssfCell.getRow().getSheet().getWorkbook());
        Configurer.Runner.doConfigure(fontConfiguration, font);

        richTextParts.add(new RichTextPart(run, font, start, end));
        return this;
    }

    @Override
    public ImageCreator png(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_PNG);
    }

    @Override
    public ImageCreator jpeg(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_JPEG);
    }

    @Override
    public ImageCreator pict(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_JPEG);
    }

    @Override
    public ImageCreator emf(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_EMF);
    }

    @Override
    public ImageCreator wmf(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_WMF);
    }

    @Override
    public ImageCreator dib(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_DIB);
    }

    private ImageCreator createImageConfigurer(int fileType) {
        return new PoiImageCreator(this, fileType);
    }

    public int getColspan() {
        if (colspan >= 1) {
            return colspan;
        }


        if (row.getSheet().getSheet().getNumMergedRegions() == 0) {
            return colspan = 1;
        }

        CellRangeAddress address = null;
        for (CellRangeAddress candidate : row.getSheet().getSheet().getMergedRegions()) {
            if (candidate.isInRange(getCell().getRowIndex(), getCell().getColumnIndex())) {
                address = candidate;
                break;
            }
        }

        if (address != null) {
            rowspan = address.getLastRow() - address.getFirstRow() + 1;
            colspan = address.getLastColumn() - address.getFirstColumn() + 1;
            return colspan;
        }

        return colspan = 1;
    }

    public int getRowspan() {
        if (rowspan >= 1) {
            return rowspan;
        }


        if (row.getSheet().getSheet().getNumMergedRegions() == 0) {
            return rowspan = 1;
        }

        CellRangeAddress address = null;
        for (CellRangeAddress candidate : row.getSheet().getSheet().getMergedRegions()) {
            if (candidate.isInRange(getCell().getRowIndex(), getCell().getColumnIndex())) {
                address = candidate;
                break;
            }
        }


        if (address != null) {
            rowspan = address.getLastRow() - address.getFirstRow() + 1;
            colspan = address.getLastColumn() - address.getFirstColumn() + 1;
            return rowspan;
        }

        return rowspan = 1;
    }

    protected XSSFCell getCell() {
        return xssfCell;
    }

    @Override
    public PoiRowDefinition getRow() {
        return row;
    }

    @Override
    public void resolve() {
        if (richTextParts != null && richTextParts.size() > 0) {
            XSSFRichTextString text = xssfCell.getRichStringCellValue();

            List<String> texts = new ArrayList<String>();

            for(RichTextPart part: richTextParts) {
                texts.add(part.getText());
            }

            text.setString(Utils.join(texts, ""));

            for (RichTextPart richTextPart : richTextParts) {
                if (richTextPart.getText() != null && richTextPart.getText().length() > 0 && richTextPart.getFont() != null) {
                    text.applyFont(richTextPart.getStart(), richTextPart.getEnd(), richTextPart.getFont().getFont());
                }
            }

            xssfCell.setCellValue(text);
        }

        if ((colspan > 1 || rowspan > 1) && poiCellStyle != null) {
            poiCellStyle.setBorderTo(getCellRangeAddress(), row.getSheet());
            // XXX: setting border messes up the style of the cell
            poiCellStyle.assignTo(this);
        }

    }

    CellRangeAddress getCellRangeAddress() {
        return new CellRangeAddress(xssfCell.getRowIndex(), xssfCell.getRowIndex() + getRowspan() - 1, xssfCell.getColumnIndex(), xssfCell.getColumnIndex() + getColspan() - 1);
    }

    @Override
    public Cell getAbove() {
        PoiRowDefinition row = this.row.getAbove();
        if (row == null) {
            return null;
        }

        PoiCellDefinition existing = (PoiCellDefinition) row.getCellByNumber(getColumn());

        if (existing != null) {
            return existing;
        }


        return createCellIfExists(getCell(row.getRow(), getColumn() - 1));
    }

    @Override
    public Cell getBellow() {
        PoiRowDefinition row = this.row.getBellow(getRowspan());
        if (row == null) {
            return null;
        }

        PoiCellDefinition existing = (PoiCellDefinition) row.getCellByNumber(getColumn());

        if (existing != null) {
            return existing;
        }


        return createCellIfExists(getCell(row.getRow(), getColumn() - 1));
    }

    @Override
    public Cell getLeft() {
        if (getColumn() == 1) {
            return null;
        }

        PoiCellDefinition existing = (PoiCellDefinition) row.getCellByNumber(getColumn() - 1);

        if (existing != null) {
            return existing;
        }


        return createCellIfExists(getCell(row.getRow(), getColumn() - 2));
    }

    @Override
    public Cell getRight() {
        if (getColumn() + getColspan() > row.getRow().getLastCellNum()) {
            return null;
        }

        PoiCellDefinition existing = (PoiCellDefinition) row.getCellByNumber(getColumn() + getColspan());

        if (existing != null) {
            return existing;
        }


        return createCellIfExists(getCell(row.getRow(), getColumn() + getColspan() - 1));
    }

    @Override
    public Cell getAboveLeft() {
        PoiRowDefinition row = this.row.getAbove();
        if (row == null) {
            return null;
        }

        if (getColumn() == 1) {
            return null;
        }

        PoiCellDefinition existing = (PoiCellDefinition) row.getCellByNumber(getColumn() - 1);

        if (existing != null) {
            return existing;
        }


        return createCellIfExists(getCell(row.getRow(), getColumn() - 2));
    }

    @Override
    public Cell getAboveRight() {
        PoiRowDefinition row = this.row.getAbove();
        if (row == null) {
            return null;
        }

        if (getColumn() == row.getRow().getLastCellNum()) {
            return null;
        }

        PoiCellDefinition existing = (PoiCellDefinition) row.getCellByNumber(getColumn() + 1);

        if (existing != null) {
            return existing;
        }


        return createCellIfExists(getCell(row.getRow(), getColumn()));
    }

    @Override
    public Cell getBellowLeft() {
        PoiRowDefinition row = this.row.getBellow();
        if (row == null) {
            return null;
        }

        if (getColumn() == 1) {
            return null;
        }

        PoiCellDefinition existing = (PoiCellDefinition) row.getCellByNumber(getColumn() - 1);

        if (existing != null) {
            return existing;
        }


        return createCellIfExists(getCell(row.getRow(), getColumn() - 2));
    }

    @Override
    public Cell getBellowRight() {
        PoiRowDefinition row = this.row.getBellow();
        if (row == null) {
            return null;
        }

        if (getColumn() == row.getRow().getLastCellNum()) {
            return null;
        }

        PoiCellDefinition existing = (PoiCellDefinition) row.getCellByNumber(getColumn() + 1);

        if (existing != null) {
            return existing;
        }


        return createCellIfExists(getCell(row.getRow(), getColumn()));
    }

    private static XSSFCell getCell(final XSSFRow row, final int column) {
        XSSFCell cell = row.getCell(column);
        if (cell != null) {
            return cell;
        }

        if (row.getSheet().getNumMergedRegions() == 0) {
            return null;
        }

        CellRangeAddress address = null;

        for (CellRangeAddress candidate: row.getSheet().getMergedRegions()) {
            if (candidate.isInRange(row.getRowNum(), column)) {
                address = candidate;
                break;
            }
        }

        return row.getSheet().getRow(address.getFirstRow()).getCell(address.getFirstColumn());
    }

    private PoiCellDefinition createCellIfExists(XSSFCell cell) {
        if (cell != null) {
            final PoiRowDefinition number = row.getSheet().getRowByNumber(cell.getRowIndex() + 1);
            return new PoiCellDefinition(number != null ? number : row.getSheet().createRowWrapper(cell.getRowIndex() + 1), cell);
        }

        return null;
    }

    @Override
    public String toString() {
        return "Cell[" + row.getSheet().getName() + "!" + getColumnAsString() + String.valueOf(row.getNumber()) + "]=" + String.valueOf(getValue());
    }

    private static void applyTo(DefaultCommentDefinition comment, XSSFCell cell) {
        if (comment.getText() == null) {
            throw new IllegalStateException("Comment text has not been set!");
        }


        XSSFWorkbook wb = cell.getRow().getSheet().getWorkbook();
        XSSFCreationHelper factory = wb.getCreationHelper();

        XSSFDrawing drawing = cell.getRow().getSheet().createDrawingPatriarch();

        XSSFClientAnchor anchor = factory.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + comment.getWidth());
        anchor.setRow1(cell.getRow().getRowNum());
        anchor.setRow2(cell.getRow().getRowNum() + comment.getHeight());

        // Create the comment and set the text+author
        XSSFComment xssfComment = drawing.createCellComment(anchor);
        xssfComment.setString(comment.getText());
        if (comment.getAuthor() != null) {
            xssfComment.setAuthor(comment.getAuthor());
        }


        // Assign the comment to the cell
        cell.setCellComment(xssfComment);
    }

    private static final double WIDTH_POINTS_PER_CM = 4.6666666666666666666667;
    private static final double WIDTH_POINTS_PER_INCH = 12;
    private static final double HEIGHT_POINTS_PER_CM = 28;
    private static final double HEIGHT_POINTS_PER_INCH = 72;
    private final PoiRowDefinition row;
    private final XSSFCell xssfCell;
    private int colspan = 0;
    private int rowspan = 0;
    private PoiCellStyleDefinition poiCellStyle;
    private List<RichTextPart> richTextParts = new ArrayList<RichTextPart>();
}
