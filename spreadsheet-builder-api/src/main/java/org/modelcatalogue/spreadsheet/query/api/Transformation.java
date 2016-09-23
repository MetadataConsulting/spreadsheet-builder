package org.modelcatalogue.spreadsheet.query.api;

public interface Transformation<S,R> {

    R transform(S source);

}
