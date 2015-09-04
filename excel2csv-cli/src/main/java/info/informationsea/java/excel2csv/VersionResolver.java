package info.informationsea.java.excel2csv;

/**
 * excel2csv
 * Copyright (C) 2015 OKAMURA Yasunobu
 * Created on 2015/08/23.
 */

import java.io.IOException;
import java.util.Properties;

public class VersionResolver {

    private static Properties getProperties() {
        Properties properties = new Properties();
        try {
            properties.load(VersionResolver.class.getResourceAsStream("/META-INF/excel2csv/version.properties"));
            return properties;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static String getGitCommit() {
        return getProperties().getProperty("git.commit");
    }

    public static String getVersion() {
        return getProperties().getProperty("version");
    }

    public static String getBuildDate() {
        return getProperties().getProperty("build.date");
    }
}
