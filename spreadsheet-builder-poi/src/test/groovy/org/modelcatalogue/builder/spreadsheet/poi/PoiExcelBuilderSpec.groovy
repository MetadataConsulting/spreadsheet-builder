package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import org.junit.Rule
import org.junit.rules.TemporaryFolder
import org.modelcatalogue.builder.spreadsheet.api.ExcelBuilder
import org.modelcatalogue.builder.spreadsheet.api.ForegroundFill
import spock.lang.Specification

import java.awt.Desktop


class PoiExcelBuilderSpec extends Specification {

    @Rule TemporaryFolder tmp

    def "create sample spreadsheet"() {
        when:
        File tmpFile = tmp.newFile("sample${System.currentTimeMillis()}.xlsx")

        ExcelBuilder builder = new PoiExcelBuilder()

        tmpFile.withOutputStream { OutputStream out ->
            builder.build(out) {
                style('zebra') {
                    font {
                        bold
                        size 12
                        color '#dddddd'
                    }
                    background yellow
                    foreground '#654321'
                    fill thinBackwardDiagonal
                }
                sheet('One') {
                    freeze 1,1
                    row {
                        cell 'First Row'
                    }

                    row()

                    row {
                        style {
                            align left center
                            border {
                                color '#abcdef'
                                style dashDotDot
                            }
                            border right, {
                                color '#00ff00'
                            }
                        }
                        cell {
                            value 'Hello'
                            name 'Salutation'
                            width auto
                        }
                        cell {
                            style 'zebra'
                            value 'World'
                            comment {
                                text 'This cell has some fancy fg/bg'
                                author 'musketyr'
                            }
                            width 50
                        }
                        cell {
                            style {
                                format 'd.m.y'
                                align center center
                            }
                            value new Date()
                            comment 'This is a date!'
                            colspan 5
                            rowspan 2
                        }
                    }
                }
                sheet('Links') {
                    freeze 1,0
                    row {
                        cell {
                            value 'Document (and a very long text)'
                            link to name 'Salutation'
                            width auto
                        }
                        cell {
                            value 'File'
                            link to file 'text.txt'
                        }
                        cell {
                            value 'URL'
                            link to url 'https://www.google.com'
                        }
                        cell {
                            value 'Mail (plain)'
                            link to email 'vladimir@orany.cz'
                        }
                        cell {
                            value 'Mail (with subject)'
                            link to email 'vladimir@orany.cz', subject: 'Testing Excel Builder', body: 'It is really great tools'
                        }
                    }
                }
                sheet ('Groups'){
                    row {
                        cell "Headline 1"
                        group {
                            cell {
                                value "Headline 2"
                                style {
                                    foreground magenta
                                    fill solidForeground
                                }
                            }
                            cell "Headline 3"
                            collapse {
                                cell "Headline 4"
                                cell "Headline 5"
                            }
                            cell "Headline 6"
                        }
                    }
                    group {
                        row {
                            cell "Some stuff"
                        }
                        collapse {
                            row {
                                cell "Something"
                            }
                            row {
                                cell "Something other"
                            }
                        }
                        row {
                            cell "Other stuff"
                        }
                    }
                }
                sheet('Fills') {
                    for(ForegroundFill foregroundFill in ForegroundFill.values()) {
                        row {
                            cell {
                                width auto
                                value foregroundFill.name()
                                style {
                                    background magenta
                                    foreground cyan
                                    fill foregroundFill
                                }
                            }
                        }
                    }
                }
            }
        }

        open tmpFile

        then:
        noExceptionThrown()
    }

    /**
     * Tries to open the file in Word. Only works locally on Mac at the moment. Ignored otherwise.
     * Main purpose of this method is to quickly open the generated file for manual review.
     * @param file file to be opened
     */
    private static void open(File file) {
        try {
            if (Desktop.desktopSupported && Desktop.desktop.isSupported(Desktop.Action.OPEN)) {
                Desktop.desktop.open(file)
                Thread.sleep(10000)
            }
        } catch(ignored) {
            // CI
        }
    }



}
