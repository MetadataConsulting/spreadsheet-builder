package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.builder.api.*;
import org.modelcatalogue.spreadsheet.impl.*;

import java.util.*;

class PoiCellDefinition extends AbstractCellDefinition {
    PoiCellDefinition(PoiRowDefinition row, XSSFCell xssfCell) {
        super(row);
        this.xssfCell = checkNotNull(xssfCell, "Cell");
    }

    private static <T> T checkNotNull(T o, String what) {
        if (o == null) {
            throw new IllegalArgumentException(what + " cannot be null");
        }

        return o;
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
    protected AbstractPendingFormula createPendingFormula(String formula) {
        return new PoiPendingFormula(this, formula);
    }

    @Override
    protected AbstractCellStyleDefinition createCellStyle() {
        return new PoiCellStyleDefinition(this);
    }

    @Override
    protected void assignStyle(CellStyleDefinition cellStyle) {
        if (cellStyle instanceof PoiCellStyleDefinition) {
            PoiCellStyleDefinition style = (PoiCellStyleDefinition) cellStyle;
            style.assignTo(this);
        } else {
            throw new IllegalArgumentException("Unsupported style: " + cellStyle);
        }
    }

    @Override
    protected void doName(String name) {
        XSSFName theName = xssfCell.getRow().getSheet().getWorkbook().createName();
        theName.setNameName(name);
        theName.setRefersToFormula(generateRefersToFormula());
    }

    @Override
    protected LinkDefinition createLinkDefinition() {
        return new PoiLinkDefinition(this.getRow().getSheet().getWorkbook(), this);
    }

    @Override
    protected FontDefinition createFontDefinition() {
        return new PoiFontDefinition(this.getRow().getSheet().getWorkbook().getWorkbook());
    }

    @Override
    protected void applyComment(DefaultCommentDefinition comment) {
        applyTo(comment, xssfCell);
    }

    private String generateRefersToFormula() {
        return "\'" + xssfCell.getSheet().getSheetName().replaceAll("'", "\\'") + "\'!" + xssfCell.getReference();
    }

    @Override
    public DimensionModifier width(double width) {
        getRow().getSheet().getSheet().setColumnWidth(xssfCell.getColumnIndex(), (int) Math.round(width * 255D));
        return new WidthModifier(this, width, WIDTH_POINTS_PER_CM, WIDTH_POINTS_PER_INCH);
    }

    @Override
    public DimensionModifier height(double height) {
        getRow().getRow().setHeightInPoints((float) height);
        return new HeightModifier(this, height, HEIGHT_POINTS_PER_CM, HEIGHT_POINTS_PER_INCH);
    }

    @Override
    public PoiCellDefinition width(Keywords.Auto auto) {
        getRow().getSheet().addAutoColumn(xssfCell.getColumnIndex());
        return this;
    }

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

    protected XSSFCell getCell() {
        return xssfCell;
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
                    text.applyFont(richTextPart.getStart(), richTextPart.getEnd(), ((PoiFontDefinition)richTextPart.getFont()).getFont());
                }
            }

            xssfCell.setCellValue(text);
        }

        if ((getColspan() > 1 || getRowspan() > 1) && cellStyle != null && cellStyle instanceof PoiCellStyleDefinition) {
            ((PoiCellStyleDefinition) cellStyle).setBorderTo(getCellRangeAddress(), getRow().getSheet());
            // XXX: setting border messes up the style of the cell
            ((PoiCellStyleDefinition) cellStyle).assignTo(this);
        }

    }

    CellRangeAddress getCellRangeAddress() {
        return new CellRangeAddress(xssfCell.getRowIndex(), xssfCell.getRowIndex() + getRowspan() - 1, xssfCell.getColumnIndex(), xssfCell.getColumnIndex() + getColspan() - 1);
    }

    @Override
    public String toString() {
        return "Cell[" + getRow().getSheet().getName() + "!" + xssfCell.getReference() + String.valueOf(getRow().getNumber()) + "]=" +xssfCell.toString();
    }

    private static void applyTo(DefaultCommentDefinition comment, XSSFCell cell) {
        if (comment.getText() == null) {
            throw new IllegalStateException("Comment text has not been set!");
        }

        XSSFComment xssfComment = cell.getCellComment();

        if (xssfComment == null) {
            XSSFWorkbook wb = cell.getRow().getSheet().getWorkbook();
            XSSFCreationHelper factory = wb.getCreationHelper();

            XSSFDrawing drawing = cell.getRow().getSheet().createDrawingPatriarch();

            XSSFClientAnchor anchor = factory.createClientAnchor();
            anchor.setCol1(cell.getColumnIndex());
            anchor.setCol2(cell.getColumnIndex() + comment.getWidth());
            anchor.setRow1(cell.getRow().getRowNum());
            anchor.setRow2(cell.getRow().getRowNum() + comment.getHeight());

            // Create the comment and set the text+author
            xssfComment = drawing.createCellComment(anchor);
        }


        XSSFRichTextString xssfCommentString = xssfComment.getString();
        if (xssfCommentString == null) {
            xssfComment.setString(comment.getText());
        } else {
            xssfCommentString.append(comment.getText());
        }

        if (comment.getAuthor() != null) {
            xssfComment.setAuthor(comment.getAuthor());
        }


        // Assign the comment to the cell
        cell.setCellComment(xssfComment);
    }

    public PoiRowDefinition getRow() {
        return (PoiRowDefinition) super.getRow();
    }

    private static final double WIDTH_POINTS_PER_CM = 4.6666666666666666666667;
    private static final double WIDTH_POINTS_PER_INCH = 12;
    private static final double HEIGHT_POINTS_PER_CM = 28;
    private static final double HEIGHT_POINTS_PER_INCH = 72;
    private final XSSFCell xssfCell;
}
