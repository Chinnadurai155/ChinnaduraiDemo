package com.rating.mobile;

import java.io.IOException;

public class runner {

    public static void main(String[] args) throws IOException {
        excelIntegration integration = new excelIntegration();
        integration.excelIntegration("src/main/resources/MobileAppsData/MobileAppsRating.xls","https://play.google.com", "https://apps.apple.com");

    }
}
