package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.ss.usermodel.ClientAnchor
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.Drawing
import org.apache.poi.ss.usermodel.Picture
import org.modelcatalogue.spreadsheet.builder.api.ImageCreator

class PoiImageCreator implements ImageCreator {

    private final PoiCellDefinition cell
    private final int type

    PoiImageCreator(PoiCellDefinition poiCell, int type) {
        this.cell = poiCell
        this.type = type
    }

    @Override
    void from(String fileOrUrl) {
        if (fileOrUrl.startsWith('https://') || fileOrUrl.startsWith('http://')) {
            from new URL(fileOrUrl).newInputStream()
            return
        }
        from new FileInputStream(new File(fileOrUrl))
    }

    @Override
    void from(InputStream stream) {
        addPicture(cell.row.sheet.sheet.workbook.addPicture(stream, type))
    }

    @Override
    void from(byte[] imageData) {
        addPicture(cell.row.sheet.sheet.workbook.addPicture(imageData, type))
    }

    void addPicture(int pictureIdx) {
        Drawing drawing = cell.row.sheet.sheet.createDrawingPatriarch();

        CreationHelper helper = cell.row.sheet.sheet.workbook.getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(cell.cell.columnIndex)
        anchor.setRow1(cell.cell.rowIndex)

        Picture pict = drawing.createPicture(anchor, pictureIdx);
        pict.resize();
    }
}
