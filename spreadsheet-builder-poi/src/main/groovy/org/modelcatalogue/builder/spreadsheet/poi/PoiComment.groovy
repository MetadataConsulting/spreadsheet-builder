package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFClientAnchor
import org.apache.poi.xssf.usermodel.XSSFComment
import org.apache.poi.xssf.usermodel.XSSFCreationHelper
import org.apache.poi.xssf.usermodel.XSSFDrawing
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.builder.spreadsheet.api.Comment

@CompileStatic class PoiComment implements Comment {

    private String author
    private String text
    private int width = DEFAULT_WIDTH
    private int height = DEFAULT_HEIGHT

    @Override
    void author(String author) {
        this.author = author
    }

    @Override
    void text(String text) {
        assert text
        this.text = text
    }

    @Override
    void width(int widthInCells) {
        assert widthInCells > 0
        this.width = widthInCells
    }

    @Override
    void height(int heightInCells) {
        assert heightInCells > 0
        this.height = heightInCells
    }

    void applyTo(XSSFCell cell) {
        if (!text) {
            throw new IllegalStateException("Comment text has not been set!")
        }
        XSSFWorkbook wb = cell.row.sheet.workbook as XSSFWorkbook
        XSSFCreationHelper factory = wb.getCreationHelper();

        XSSFDrawing drawing = cell.row.sheet.createDrawingPatriarch() as XSSFDrawing;

        XSSFClientAnchor anchor = factory.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + width);
        anchor.setRow1(cell.row.getRowNum());
        anchor.setRow2(cell.row.getRowNum() + height);

        // Create the comment and set the text+author
        XSSFComment comment = drawing.createCellComment(anchor);
        comment.setString(text)
        if (author) {
            comment.setAuthor(author)
        }

        // Assign the comment to the cell
        cell.setCellComment(comment)


    }
}
