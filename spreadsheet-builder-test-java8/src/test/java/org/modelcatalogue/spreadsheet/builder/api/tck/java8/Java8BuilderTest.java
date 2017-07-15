package org.modelcatalogue.spreadsheet.builder.api.tck.java8;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.builder.api.CellDefinition;
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder;
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition;
import org.modelcatalogue.spreadsheet.builder.poi.PoiSpreadsheetBuilder;
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteria;
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteriaResult;
import org.modelcatalogue.spreadsheet.query.poi.PoiSpreadsheetCriteria;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import static org.modelcatalogue.spreadsheet.api.Keywords.*;
import static org.modelcatalogue.spreadsheet.api.Color.*;

public class Java8BuilderTest {

    @Rule public TemporaryFolder tmp = new TemporaryFolder();

    @Test public void testBuilder() throws IOException {
        File excel = tmp.newFile();
        PoiSpreadsheetBuilder.INSTANCE.build(w ->
            w.sheet("test", s ->
                s.row(r ->
                    r.cell(c ->
                        c.value("Hello World")
                    )
                )
            )
        ).writeTo(excel);

        assertEquals(1, PoiSpreadsheetCriteria.FACTORY.forFile(excel).query(w ->
            w.sheet("test", s ->
               s.row(r ->
                   r.cell(c ->
                       c.value("Hello World")
                   )
               )
            )
        ).getCells().size());
    }

    @Test public void testBuilderFull() throws IOException {
        Date today = new Date();
        File excel = tmp.newFile();

        buildSpreadsheet(PoiSpreadsheetBuilder.INSTANCE, today).writeTo(excel);

        SpreadsheetCriteria matcher = PoiSpreadsheetCriteria.FACTORY.forFile(excel);

        SpreadsheetCriteriaResult allCells = matcher.all();

        assertEquals(80130, allCells.getCells().size());
        assertEquals(21, allCells.getSheets().size());
        assertEquals(20065, allCells.getRows().size());


        SpreadsheetCriteriaResult sampleCells = matcher.query(w -> w.sheet("Sample"));

        assertEquals(2, sampleCells.getCells().size());
        assertEquals(1, sampleCells.getSheets().size());
        assertEquals(1, sampleCells.getRows().size());

        SpreadsheetCriteriaResult rowCells = matcher.query(w ->
            w.sheet("many rows", s ->
                    s.row(1)
            )
        );

        assertEquals(4, rowCells.getCells().size());
        assertEquals(1, rowCells.getSheets().size());
        assertEquals(1, rowCells.getRows().size());

        Row manyRowsHeader = matcher.query(w ->
            w.sheet("many rows", s ->
                s.row(1)
            )
        ).getRow();

        assertNotNull(manyRowsHeader);

        Row manyRowsDataRow = matcher.query(w -> w.sheet("many rows", s -> s.row(2))).getRow();

        DataRow dataRow = DataRow.create(manyRowsDataRow, manyRowsHeader);

        assertNotNull(dataRow.get("One"));
        assertEquals("1", dataRow.get("One").getValue());


        Map<String, Integer> dataRowMapping = new HashMap<>();
        dataRowMapping.put("primo", 1);


        DataRow dataRowFromMapping = DataRow.create(dataRowMapping, manyRowsDataRow);
        assertNotNull(dataRowFromMapping.get("primo"));
        assertEquals("1", dataRowFromMapping.get("primo").getValue());


        SpreadsheetCriteriaResult someCells = matcher.query(w ->
            w.sheet(s ->
               s.row(r ->
                   r.cell(c ->
                      c.date(today)
                   )
               )
            )
        );

        assertEquals(1, someCells.getCells().size());


        SpreadsheetCriteriaResult commentedCells = matcher.query(w ->
            w.sheet(s ->
                s.row(r ->
                    r.cell(c ->
                        c.comment("This is a date!")
                    )
                )
            )
        );

        assertEquals(1, commentedCells.getCells().size());


        SpreadsheetCriteriaResult namedCells = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.name("_Cell10");
                    });
                });
            });
        });

        assertEquals(1, namedCells.getCells().size());


        SpreadsheetCriteriaResult dateCells = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.style(st -> {
                            st.format("d.m.y");
                        });
                    });
                });
            });
        });

        assertEquals(1, dateCells.getCells().size());

        SpreadsheetCriteriaResult filledCells = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.style(st -> {
                            st.fill(fineDots);
                        });
                    });
                });
            });
        });

        assertEquals(1, filledCells.getCells().size());

        SpreadsheetCriteriaResult magentaCells = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.style(st -> {
                            st.foreground(aquamarine);
                        });
                    });
                });
            });
        });

        assertEquals(1, magentaCells.getCells().size());

        SpreadsheetCriteriaResult redOnes = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.style(st -> {
                            st.font(f -> {
                                f.color(red);
                            });
                        });
                    });
                });
            });
        });

        assertEquals(20006, redOnes.getCells().size());
        assertEquals(20004, redOnes.getRows().size());


        SpreadsheetCriteriaResult boldOnes = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.style(st -> {
                            st.font(f -> {
                                f.make(bold);
                            });
                        });
                    });
                });
            });
        });

        assertEquals(5, boldOnes.getCells().size());

        SpreadsheetCriteriaResult bigOnes = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.style(st -> {
                            st.font(f -> {
                                f.size(22);
                            });
                        });
                    });
                });
            });
        });

        assertEquals(40002, bigOnes.getCells().size());

        SpreadsheetCriteriaResult bordered = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.style(st -> {
                            st.border(top, b -> {
                                b.style(thin);
                            });
                        });
                    });
                });
            });
        });

        assertEquals(10, bordered.getCells().size());

        SpreadsheetCriteriaResult combined = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("Bold Red 22");
                        c.style(st -> {
                            st.font(f -> {
                                f.color(red);
                            });
                        });
                    });
                });
            });
        });

        assertEquals(1, combined.getCells().size());

        SpreadsheetCriteriaResult conjunction = matcher.query(w -> {
            w.sheet(s -> {
                s.row(r -> {
                    r.or(o -> {
                        o.cell(c -> {
                            c.value("Bold Red 22");
                        });
                        o.cell(c -> {
                            c.value("A");
                        });
                    });
                });
            });
        });

        assertEquals(3, conjunction.getCells().size());

        SpreadsheetCriteriaResult traversal = matcher.query(w -> {
            w.sheet("Traversal", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("E");
                    });
                });
            });
        });

        assertEquals(1, traversal.getCells().size());

        Cell cellE = traversal.iterator().next();

        assertEquals("Traversal", cellE.getRow().getSheet().getName());
        assertNotNull(cellE.getRow().getSheet().getPrevious());
        assertEquals("Formula", cellE.getRow().getSheet().getPrevious().getName());
        assertNotNull(cellE.getRow().getSheet().getNext());
        assertEquals("Border", cellE.getRow().getSheet().getNext().getName());
        assertEquals(2, cellE.getRow().getNumber());
        assertNotNull(cellE.getRow().getAbove());
        assertEquals(1, cellE.getRow().getAbove().getNumber());
        assertNotNull(cellE.getRow().getBellow());
        assertEquals(3, cellE.getRow().getBellow().getNumber());
        assertEquals(2, cellE.getColspan());
        assertNotNull(cellE.getAboveLeft());
        assertEquals("A", cellE.getAboveLeft().getValue());
        assertNotNull(cellE.getAbove());
        assertEquals("B", cellE.getAbove().getValue());
        assertNotNull(cellE.getAboveRight());
        assertEquals("C", cellE.getAboveRight().getValue());
        assertNotNull(cellE.getLeft());
        assertEquals("D", cellE.getLeft().getValue());
        assertNotNull(cellE.getRight());
        assertEquals("F", cellE.getRight().getValue());
        assertNotNull(cellE.getBellowLeft());
        assertEquals("G", cellE.getBellowLeft().getValue());
        assertNotNull(cellE.getBellowRight());
        assertEquals("I", cellE.getBellowRight().getValue());
        assertNotNull(cellE.getBellow());
        assertEquals("H", cellE.getBellow().getValue());
        assertEquals("J", cellE.getBellow().getBellow().getValue());

        SpreadsheetCriteriaResult zeroCells = matcher.query(w -> {
            w.sheet("Zero", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value(0);
                    });
                });
            });
        });

        assertEquals(1, zeroCells.getCells().size());
        assertEquals(0d, zeroCells.getCell().getValue());

        Cell noneCell = matcher.find(w -> {
            w.sheet("Styles", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("NONE");
                    });
                });
            });
        });

        Cell redCell = matcher.find(w -> {
            w.sheet("Styles", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("RED");
                    });
                });
            });
        });

        Cell blueCell = matcher.find(w -> {
            w.sheet("Styles", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("BLUE");
                    });
                });
            });
        });

        Cell greenCell = matcher.find(w -> {
            w.sheet("Styles", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("GREEN");
                    });
                });
            });
        });

        assertNull(noneCell.getStyle().getForeground());
        assertEquals(red, redCell.getStyle().getForeground());
        assertEquals(blue, blueCell.getStyle().getForeground());
        assertEquals(green, greenCell.getStyle().getForeground());

        assertEquals(1, matcher.query(w -> {
            w.sheet(s -> {
                s.page(p -> {
                    p.paper(Keywords.Paper.A5);
                    p.orientation(Keywords.Orientation.LANDSCAPE);
                });
            });
        }).getCells().size());
        assertEquals(1, matcher.query(w -> {
            w.sheet("Hidden", s -> {
                s.hide();
            });
        }).getSheets().size());
        assertEquals(1, matcher.query(w -> {
            w.sheet("Very hidden", s -> {
                s.hideCompletely();
            });
        }).getSheets().size());

        SpreadsheetDefinition definition = PoiSpreadsheetBuilder.INSTANCE.build(excel, w -> {
            w.sheet("Sample", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("Ahoj");
                    });
                    r.cell(c -> {
                        c.value("Svete");
                    });
                });
            });
        });


        File tmpFile = tmp.newFile();
        definition.writeTo(tmpFile);

        assertEquals(0, PoiSpreadsheetCriteria.FACTORY.forFile(tmpFile).query(w -> {
            w.sheet("Sample", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("Hello");
                    });

                });
            });
        }).getCells().size());


        assertEquals(1, PoiSpreadsheetCriteria.FACTORY.forFile(tmpFile).query(w -> {
            w.sheet("Sample", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("Ahoj");
                    });

                });
            });
        }).getCells().size());
    }

    private static SpreadsheetDefinition buildSpreadsheet(SpreadsheetBuilder builder, Date today) {
        return builder.build(w -> {
            w.style("red", s -> {
                s.font( f -> {
                    f.color(red);
                });
            });
            w.style("bold", s -> {
                s.font(f -> {
                    f.make(bold);
                });
            });

            w.style("borders", s -> {
                s.font(f -> {
                    f.color(red);
                });
                s.border(b -> {
                    b.style(thin).color(black);
                });
            });

            w.style("centered", s -> {
                    s.align(center, center);
            });

            w.style("redfg", s -> {
                    s.foreground(red);
            });

            w.style("greenfg", s -> {
                    s.foreground(green);
            });

            w.style("bluefg", s -> {
                    s.foreground(blue);
            });

            w.style("nonefg", s -> {});


            w.apply(MyStyles.class);

            w.sheet("many rows", s -> {
                s.filter(auto);
                s.row(r -> {
                    r.cell("One");
                    r.cell("Two");
                    r.cell("Three");
                    r.cell("Four");
                });
                for (int i = 0; i < 20000; i++) {
                    s.row(r -> {
                        r.cell(c -> {
                            c.value("1");
                            c.styles("h1", "red");
                        });
                        r.cell( c -> {
                            c.value("2");
                            c.style("h2");
                        });
                        r.cell( c -> {
                            c.value("3");
                            c.style("h1");
                        });
                        r.cell( c -> {
                            c.value("4");
                            c.style("h2");
                        });
                    });
                }
            });

            w.sheet("Sample", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("Hello");
                        c.style("h1");
                    });
                    r.cell(c -> {
                        c.value("World");
                        c.style("h2");
                    });
                });
            });

            w.sheet("Fonts", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.width(auto);
                        c.value("Bold Red 22");
                        c.style(st -> {
                            st.font(f -> {
                                f.make(bold);
                                f.color(red);
                                f.size(22);
                            });
                        });
                    });
                    r.cell(c -> {
                        c.width(auto);
                        c.value("Underline Courier New");
                        c.style(st -> {
                            st.font(f -> {
                                f.make(underline);
                                f.name("Courier New");
                            });
                        });
                    });
                    r.cell(c -> {
                        c.width(auto);
                        c.value("Italic");
                        c.style(st -> {
                            st.font(f -> {
                                f.make(italic);
                            });
                        });
                    });
                    r.cell(c -> {
                        c.width(auto);
                        c.value("Strikeout");
                        c.style(st -> {
                            st.font(f -> {
                                f.make(strikeout);
                            });
                        });
                    });
                });
            });

            w.sheet("Image", s -> {
                s.row(3, r -> {
                    r.cell("C", c -> {
//                         c.png(c.getImage()).from("https://goo.gl/UcL1wy");
                    });
                });
            });

            w.sheet("Rich Text", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.text("Little");
                        c.text(" ");
                        c.text("Red", f -> {
                                f.color(red);
                                f.size(22);
                                f.name("Times New Roman");
                        });
                        c.text(" ");
                        c.text("Riding", f -> {
                                f.make(italic);
                                f.size(18);
                        });
                        c.text(" ");
                        c.text("Hood", f -> {
                                f.make(bold);
                                f.size(22);
                        });

                    });
                });
                s.row(r -> {
                    r.cell(c -> {
                        c.style(st -> {
                            st.wrap(text);
                        });
                        c.text("First Line");
                        c.text("\n");
                        c.text("Second Line", f -> {
                                f.make(bold);
                                f.size(12);
                        });
                        c.text("\n");

                        Map<String, String> values = new LinkedHashMap<>();
                        values.put("foo","bar");
                        values.put("boo", "cow");
                        values.put("empty", "");
                        values.put("", "nothing");

                        for (Map.Entry<String, String> entry : values.entrySet()) {
                            printMap(c, entry);
                        }
                        c.text("\n\n");
                        c.text("Next line after two spaces");
                        c.text("\n");
                        c.text("Last line", f -> {
                                f.make(italic);
                        });
                        c.text("\n");
                    });
                });

                s.row(r -> {
                    r.cell(c -> {
                        c.text("Genomics England Consent Withdrawal Options");
                        c.text("\n\n");
                        c.text("Enumerations", f -> {
                                f.size(12);
                                f.make(bold);
                        });
                        c.text("\n");

                        Map<String, String> withdrawals = new LinkedHashMap<>();
                        withdrawals.put("FULL_WITHDRAWAL", "OPTION 2: FULL WITHDRAWAL: No further use");
                        withdrawals.put("PARTIAL_WITHDRAWAL", "OPTION 1: PARTIAL WITHDRAWAL: No further contact");

                        for (Map.Entry<String, String> entry : withdrawals.entrySet()) {
                            printMap(c, entry);
                        }
                    });
                });
            });

            w.sheet("Cell Addressing", s -> {
                s.filter(auto);
                s.row(2, r -> {
                    r.style(st -> {
                        st.background(whiteSmoke);
                        st.border(top, bottom, b -> {
                            b.style(thin);
                            b.color(black);
                        });
                    });
                    r.cell("B", c -> {
                        c.value("A");
                        c.style(st -> {
                            st.border(left,  b -> {
                                b.style(thin);
                                b.color(black);
                            });
                        });
                    });
                    r.cell("B");
                    r.cell(c -> {
                        c.value("C");
                        c.style(st -> {
                            st.border(right, b -> {
                                b.style(thin);
                                b.color(black);
                            });
                        });
                    });
                });
                s.row(r -> {
                    r.cell("B", c -> c.value(1));
                    r.cell(2);
                    r.cell(3);
                });
            });
            w.sheet("One", s -> {
                s.freeze(1, 1);
                s.row(r -> {
                    r.cell("First Row");
                });

                s.row(r -> {
                    r.cell("AC", c -> c.value("AC"));
                    r.cell("BE", c -> c.value("BE"));
                });

                s.row(r -> {
                    r.style(st -> {
                        st.align(center, left);
                        st.border(b -> {
                            b.color("#abcdef");
                            b.style(dashDotDot);
                        });
                        st.border(right, b -> {
                            b.color("#00ff00");
                        });
                    });
                    r.cell(c -> {
                        c.value("Hello");
                        c.name("Salutation");
                        c.width(auto);
                    });
                    r.cell(c -> {
                        c.style("bold");
                        c.value("World");
                        c.comment(comm -> comm.text("This cell has some fancy fg/bg").author("musketyr"));
                        c.width(50);
                    });
                    r.cell(c -> {
                        c.style(st -> {
                            st.format("d.m.y");
                            st.align(center, center);
                        });
                        c.value(today);
                        c.comment("This is a date!");
                        c.colspan(5);
                        c.rowspan(2);
                    });
                });
            });
            w.sheet("Links", s -> {
                s.freeze(1, 0);
                s.row(r -> {
                    r.cell(c -> {
                        c.value("Document (and a very long text)");
                        c.link(to).name("Salutation");
                        c.width(auto);
                    });
                    r.cell(c -> {
                        c.value("File");
                        c.link(to).file("text.txt");
                    });
                    r.cell(c -> {
                        c.value("URL");
                        c.link(to).url("https://www.google.com");
                    });
                    r.cell(c -> {
                        c.value("Mail (plain)");
                        c.link(to).email("vladimir@orany.cz");
                    });
                    r.cell(c -> {
                        c.value("Mail (with subject)");
                        Map<String, String> email = new LinkedHashMap<>();
                        email.put("subject", "Testing Excel Builder");
                        email.put("body", "It is really great tools");
                        c.link(to).email(email,"vladimir@orany.cz");
                    });
                });
            });
            w.sheet("Groups", s -> {
                s.row(r -> {
                    r.cell("Headline 1");
                    r.group(g -> {
                        g.cell(c -> {
                            c.value("Headline 2");
                            c.style(st -> {
                                st.foreground(aquamarine);
                                st.fill(solidForeground);
                            });
                        });
                        g.cell("Headline 3");
                        g.collapse(clps -> {
                            clps.cell("Headline 4");
                            clps.cell("Headline 5");
                        });
                        g.cell("Headline 6");
                    });
                });
                s.group(g -> {
                    g.row(r -> {
                        r.cell("Some stuff");
                    });
                    g.collapse(clps -> {
                        clps.row(r -> {
                            r.cell("Something");
                        });
                        clps.row(r -> {
                            r.cell("Something other");
                        });
                    });
                    s.row(r -> {
                        r.cell("Other stuff");
                    });
                });
            });
            w.sheet("Fills", s -> {
                for (ForegroundFill foregroundFill : ForegroundFill.values()) {
                    s.row(r -> {
                        r.cell(c -> {
                            c.width(auto);
                            c.value(foregroundFill.name());
                        });
                        r.cell(c -> {
                            c.style(st -> {
                                st.background("#FF8C00"); // darkOrange
                                st.foreground(brown);
                                st.fill(foregroundFill);
                            });
                        });
                    });
                }
            });
            w.sheet("Formula", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value(10);
                        c.name("_Cell10");
                    });
                    r.cell(c -> {
                        c.value(20);
                        c.name("_Cell20");
                    });
                    r.cell(c -> {
                        c.formula("SUM(#{_Cell10}:#{_Cell20})");
                    });
                });
            });
            // create sheet with same name, it should query the
            w.sheet("Formula", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value(30);
                        c.name("_Cell30");
                    });
                    r.cell(c -> {
                        c.value(40);
                        c.name("_Cell40");
                    });
                    r.cell(c -> {
                        c.formula("SUM(#{_Cell30}:#{_Cell40})");
                    });
                });
            });
            w.sheet("Traversal", s -> {
                s.row(r -> {
                    r.cell("A");
                    r.cell("B");
                    r.cell("C");
                });
                s.row(r -> {
                    r.cell("D");
                    r.cell(c -> {
                        c.value("E");
                        c.colspan(2);
                    });
                    r.cell("F");
                });
                s.row(r -> {
                    r.cell("G");
                    r.cell(c -> {
                        c.value("H");
                        c.rowspan(2);
                    });
                    r.cell("I");
                });
                s.row(5, r -> {
                    r.cell("B", c -> {
                        c.value("J");
                    });
                });
            });
            w.sheet("Border", s -> {
                s.row(r -> {
                    r.style("borders");
                    r.cell(c -> {
                        c.value(1);
                        c.colspan(2);
                    });
                    r.cell(c -> {
                        c.value(2);
                    });
                });
                s.row(r -> {
                    r.style("borders");
                    r.cell(c -> {
                        c.value(1);
                        c.colspan(2);
                    });
                    r.cell(c -> {
                        c.value(2);
                    });

                });
            });
            w.sheet("Zero", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value(0);
                    });
                });
            });
            w.sheet("Filtered", s -> {
                s.filter(auto);
                s.row(r -> {
                    r.cell("Name");
                    r.cell("Profession");
                });
                s.row(r -> {
                    r.cell("Donald");
                    r.cell("Sailor");
                });
                s.row(r -> {
                    r.cell("Bob");
                    r.cell("Builder");
                });
            });

            w.sheet("Styles", s -> {
                s.row(r -> {
                    r.style("nonefg");
                    r.cell(c -> {
                        c.value("NONE");
                        c.rowspan(3);
                        c.style("centered");
                    });
                });
                s.row(4, r -> {
                    r.style("redfg");
                    r.cell(c -> {
                        c.value("RED");
                        c.rowspan(3);
                        c.style("centered");
                    });
                });
                s.row(7, r -> {
                    r.style("greenfg");
                    r.cell(c -> {
                        c.value("GREEN");
                        c.rowspan(3);
                        c.styles("centered", "bold");
                    });
                });
                s.row(10, r -> {
                    r.styles("bluefg", "borders");
                    r.cell(c -> {
                        c.value("BLUE");
                        c.rowspan(3);
                        c.styles ("centered", "bold");
                    });
                });
            });

            w.sheet("Dimensions", s -> {
                s.row(r -> {
                    r.cell(c -> {
                        c.value("cm");
                        c.height(1).cm();
                        c.width(1).cm();
                    });
                });

                s.row(r -> {
                    r.cell("B", c -> {
                        c.value("inches");
                        c.width(1).inch();
                        c.height(1).inch();
                    });
                });
                s.row(r -> {
                    r.cell("C", c -> {
                        c.value("points");
                        c.width(10);
                        c.height(50);
                    });
                });
            });

            w.sheet("Custom Page", s -> {
                s.page( page -> {
                    page.paper(A5);
                    page.fit(width).to(1);
                    page.orientation(landscape);
                });
                s.row(r -> {
                    r.cell("A5 Landcapse");
                });
            });
            w.sheet("Broken row styles", s ->
                    s.row(r ->
                            r.styles("bold", "redfg").cell(c ->
                                c.value("bold and RED")
                            )
                    )
            );
            w.sheet("Hidden", s ->
                    s.hide()
            );
            w.sheet("Very hidden", s ->
                    s.hideCompletely()
            );
        });
    }

    private static void printMap(CellDefinition c, Map.Entry<String, String> entry) {
        c.text(entry.getKey(), f -> {
                f.make(bold);
        });
        c.text(": ");
        c.text(entry.getValue());
        c.text("\n");
    }

}
