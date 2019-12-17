package com.excelcounter.utils;

import javax.swing.*;
import java.io.File;

public class Utils {
    private static String lastDir = null;

    public static JFileChooser getFileChooser() {
        if (lastDir != null) {
            return new JFileChooser(lastDir);
        } else {
            return new JFileChooser();
        }
    }

    public static void setLastDir(File file) {
        lastDir = file.getParent();
    }
}