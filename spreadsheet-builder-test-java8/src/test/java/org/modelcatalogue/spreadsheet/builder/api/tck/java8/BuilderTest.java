package org.modelcatalogue.spreadsheet.builder.api.tck.java8;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.modelcatalogue.spreadsheet.builder.poi.PoiSpreadsheetBuilder;
import org.modelcatalogue.spreadsheet.query.poi.PoiSpreadsheetCriteria;

import java.io.File;
import java.io.IOException;

import static org.junit.Assert.assertEquals;

public class BuilderTest {

    @Rule public TemporaryFolder tmp = new TemporaryFolder();

    @Test public void testBuilder() throws IOException {
        File excel = tmp.newFile();
        PoiSpreadsheetBuilder.INSTANCE.build((w) -> {
            w.sheet("test", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("Hello World");
                    });
                });
            });
        }).writeTo(excel);

        assertEquals(1, PoiSpreadsheetCriteria.FACTORY.forFile(excel).query(w -> {
            w.sheet(s -> {
               s.row(r -> {
                   r.cell(c -> {
                       c.value("Hello World");
                   });
               });
            });
        }).getCells().size());
    }

}
