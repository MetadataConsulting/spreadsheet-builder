package builders.dsl.spreadsheet.builder.api;

import java.io.InputStream;

public interface ImageCreator {

    CellDefinition from(String fileOrUrl);
    CellDefinition from(InputStream stream);
    CellDefinition from(byte[] imageData);

}
