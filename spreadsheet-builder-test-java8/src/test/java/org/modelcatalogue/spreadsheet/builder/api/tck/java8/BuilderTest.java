package org.modelcatalogue.spreadsheet.builder.api.tck.java8;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.modelcatalogue.spreadsheet.api.ForegroundFill;
import org.modelcatalogue.spreadsheet.builder.api.CellDefinition;
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder;
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition;
import org.modelcatalogue.spreadsheet.builder.poi.PoiSpreadsheetBuilder;
import org.modelcatalogue.spreadsheet.query.poi.PoiSpreadsheetCriteria;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.modelcatalogue.spreadsheet.api.BorderStyle.DASH_DOT_DOT;
import static org.modelcatalogue.spreadsheet.api.BorderStyle.THIN;
import static org.modelcatalogue.spreadsheet.api.FontStyle.*;
import static org.modelcatalogue.spreadsheet.api.ForegroundFill.SOLID_FOREGROUND;
import static org.modelcatalogue.spreadsheet.api.HTMLColorProvider.*;
import static org.modelcatalogue.spreadsheet.api.Keywords.Auto.AUTO;
import static org.modelcatalogue.spreadsheet.api.Keywords.BorderSide.*;

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
            w.sheet("test", s -> {
               s.row(r -> {
                   r.cell(c -> {
                       c.value("Hello World");
                   });
               });
            });
        }).getCells().size());
    }

    @Test public void testBuilderFull() throws IOException {
        File excel = tmp.newFile();
        buildSpreadsheet(PoiSpreadsheetBuilder.INSTANCE, new Date()).writeTo(excel);
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
                    f.make(BOLD);
                });
            });

            w.style("h1", s -> {
                s.font(f -> {
                    f.make(BOLD);
                });
            });

            w.style("h2", s -> {
                s.font(f -> {
                    f.make(BOLD);
                });
            });

            w.style("borders", s -> {
                s.font(f -> {
                    f.color(red);
                });
                s.border(b -> {
                    b.style(THIN).color(black);
                });
            });

            w.style("centered", s -> {
                    s.align(s.getCenter()).getCenter();
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
                s.filter(s.getAuto());
                s.row(r -> {
                    r.cell("One");
                    r.cell("Two");
                    r.cell("Three");
                    r.cell("Four");
                });
                for (int i = 0; i < 2000; i++) {
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
                        c.width(c.getAuto());
                        c.value("Bold Red 22");
                        c.style(st -> {
                            st.font(f -> {
                                f.make(BOLD);
                                f.color(red);
                                f.size(22);
                            });
                        });
                    });
                    r.cell(c -> {
                        c.width(c.getAuto());
                        c.value("Underline Courier New");
                        c.style(st -> {
                            st.font(f -> {
                                f.make(UNDERLINE);
                                f.name("Courier New");
                            });
                        });
                    });
                    r.cell(c -> {
                        c.width(c.getAuto());
                        c.value("Italic");
                        c.style(st -> {
                            st.font(f -> {
                                f.make(ITALIC);
                            });
                        });
                    });
                    r.cell(c -> {
                        c.width(c.getAuto());
                        c.value("Strikeout");
                        c.style(st -> {
                            st.font(f -> {
                                f.make(STRIKEOUT);
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
                                f.make(ITALIC);
                                f.size(18);
                        });
                        c.text(" ");
                        c.text("Hood", f -> {
                                f.make(BOLD);
                                f.size(22);
                        });

                    });
                });
                s.row(r -> {
                    r.cell(c -> {
                        c.style(st -> {
                            st.wrap(st.getText());
                        });
                        c.text("First Line");
                        c.text("\n");
                        c.text("Second Line", f -> {
                                f.make(BOLD);
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
                                f.make(ITALIC);
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
                                f.make(BOLD);
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
                s.filter(s.getAuto());
                s.row(2, r -> {
                    r.style(st -> {
                        st.background(whiteSmoke);
                        st.border(TOP, BOTTOM, b -> {
                            b.style(THIN);
                            b.color(black);
                        });
                    });
                    r.cell("B", c -> {
                        c.value("A");
                        c.style(st -> {
                            st.border(LEFT,  b -> {
                                b.style(THIN);
                                b.color(black);
                            });
                        });
                    });
                    r.cell("B");
                    r.cell(c -> {
                        c.value("C");
                        c.style(st -> {
                            st.border(RIGHT, b -> {
                                b.style(THIN);
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
                        st.align(st.getCenter()).getLeft();
                        st.border(b -> {
                            b.color("#abcdef");
                            b.style(DASH_DOT_DOT);
                        });
                        st.border(RIGHT, b -> {
                            b.color("#00ff00");
                        });
                    });
                    r.cell(c -> {
                        c.value("Hello");
                        c.name("Salutation");
                        c.width(c.getAuto());
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
                            st.align(st.getCenter()).getCenter();
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
                        c.link(c.getTo()).name("Salutation");
                        c.width(c.getAuto());
                    });
                    r.cell(c -> {
                        c.value("File");
                        c.link(c.getTo()).file("text.txt");
                    });
                    r.cell(c -> {
                        c.value("URL");
                        c.link(c.getTo()).url("https://www.google.com");
                    });
                    r.cell(c -> {
                        c.value("Mail (plain)");
                        c.link(c.getTo()).email("vladimir@orany.cz");
                    });
                    r.cell(c -> {
                        c.value("Mail (with subject)");
                        Map<String, String> email = new LinkedHashMap<>();
                        email.put("subject", "Testing Excel Builder");
                        email.put("body", "It is really great tools");
                        c.link(c.getTo()).email(email,"vladimir@orany.cz");
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
                                st.fill(SOLID_FOREGROUND);
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
                            c.width(c.getAuto());
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
                s.filter(AUTO);
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
                        c.height(1).getCm();
                        c.width(1).getCm();
                    });
                });

                s.row(r -> {
                    r.cell("B", c -> {
                        c.value("inches");
                        c.width(1).getInch();
                        c.height(1).getInch();
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
                    page.paper(page.getA5());
                    page.fit(page.getWidth()).to(1);
                    page.orientation(page.getLandscape());
                });
                s.row(r -> {
                    r.cell("A5 Landcapse");
                });
            });
            w.sheet("Broken row styles", s -> {
                s.row(r -> {
                    r.styles("bold", "redfg");
                    r.cell(c -> {
                        c.value("BOLD and RED");
                    });
                });
            });
        });
    }

    private static void printMap(CellDefinition c, Map.Entry<String, String> entry) {
        c.text(entry.getKey(), f -> {
                f.make(BOLD);
        });
        c.text(": ");
        c.text(entry.getValue());
        c.text("\n");
    }

}
