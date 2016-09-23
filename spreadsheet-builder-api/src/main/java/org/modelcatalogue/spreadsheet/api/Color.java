package org.modelcatalogue.spreadsheet.api;

import java.util.Arrays;

public final class Color {

    private final String hex;

    public Color(String hex) {
        if (!hex.matches("#[\\dA-Fa-f]{6}")) {
            throw new IllegalArgumentException("Wrong format for color: " + hex);
        }
        this.hex = hex.toUpperCase();
    }

    public Color(byte[] rgb) {
        if (rgb.length != 3) {
            throw new IllegalArgumentException("Wrong number of parts in: " + Arrays.toString(rgb));
        }
        this.hex = String.format("#%02X%02X%02X", rgb[0], rgb[1], rgb[2]);
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

    @Override
    public String toString() {
        return "Color[" + getHex() + "]";
    }
}
