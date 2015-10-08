package org.modelcatalogue.builder.spreadsheet.api;

public class Color {

    private final String hex;

    public Color(String hex) {
        this.hex = hex;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }

        Color color = (Color) o;

        return !(hex != null ? !hex.equals(color.hex) : color.hex != null);

    }

    @Override
    public int hashCode() {
        return hex != null ? hex.hashCode() : 0;
    }

    public String getHex() {
        return hex;
    }
}
