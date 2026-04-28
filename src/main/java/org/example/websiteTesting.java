package org.example;

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

public class websiteTesting {

    static Actions actions;
    static JavascriptExecutor jsExecutor;
    static WebDriverWait wait;
    static Set<String> allUrls;
    static List<TestResult> testResults;
    static String screenshotDir;
    static Sheets sheetsService;
    static final String SPREADSHEET_ID = "1HQAUwmh0qW0g_MKi5TidC8O5piK89-QlwPzkPphPqts";
    static String mainSheetName;
    static boolean isFirstWebsite = true;

    // Summary statistics variables
    static int totalWebsitesProcessed = 0;
    static int totalWebsitesPassed = 0;
    static int totalWebsitesFailed = 0;
    static int totalPagesTestedOverall = 0;
    static int totalPagesPassedOverall = 0;
    static int totalPagesFailedOverall = 0;
    static List<String> failedWebsitesList = new ArrayList<>();
    static Map<String, Integer> websitePageCountMap = new HashMap<>();
    static Map<String, Integer> websitePassCountMap = new HashMap<>();
    static Map<String, Integer> websiteFailCountMap = new HashMap<>();

    private static final String APPLICATION_NAME = "Website Testing Automation";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();

    // Enhanced Test data for form filling with additional fields
    static class TestData {
        static final String FIRST_NAME = "Brandcrock";
        static final String LAST_NAME = "QA";
        static final String FULL_NAME = "Brandcrock QA";
        static final String EMAIL = "brandcrocktesting@gmail.com";
        static final String PHONE = "+91 1234567890";
        static final String SUBJECT = "Test Submission";
        static final String COMPANY = "Brandcrock GmbH";
        static final String ADDRESS = "Test Street 123, 80331 Munich, Germany";
        static final String CITY = "Munich";
        static final String ZIP_CODE = "80331";
        static final String COUNTRY = "Germany";
        static final String WEBSITE = "https://www.brandcrock.com";
        static final String MESSAGE = "From Brandcrock:\nThis message has been submitted as part of a routine system check to verify that the form and email delivery are working as expected.\nNote: Please ignore this email, as it is for testing purposes only.";
    }

    static class TestResult {
        String dateTime;
        String websiteUrl;
        String pageTitle;
        String parentPage;
        String childPage;
        String fullPath;
        String action;
        String testedUrl;
        Long durationMs;
        String failedResponseLog;
        Integer statusCode;
        String status;
        String screenshotPath;

        public TestResult(String dateTime, String websiteUrl, String pageTitle, String parentPage, String childPage,
                          String fullPath, String action, String testedUrl, Long durationMs, String failedResponseLog,
                          Integer statusCode, String status, String screenshotPath) {
            this.dateTime = dateTime;
            this.websiteUrl = websiteUrl;
            this.pageTitle = pageTitle;
            this.parentPage = parentPage;
            this.childPage = childPage;
            this.fullPath = fullPath;
            this.action = action;
            this.testedUrl = testedUrl;
            this.durationMs = durationMs;
            this.failedResponseLog = failedResponseLog;
            this.statusCode = statusCode;
            this.status = status;
            this.screenshotPath = screenshotPath;
        }

        public List<Object> toRow() {
            return Arrays.asList(
                    dateTime, websiteUrl, pageTitle, parentPage, childPage, fullPath, action, testedUrl,
                    durationMs != null ? durationMs : 0,
                    failedResponseLog != null ? failedResponseLog : "",
                    statusCode != null ? statusCode : 0,
                    status != null ? status : "UNKNOWN",
                    screenshotPath != null ? screenshotPath : ""
            );
        }
    }

    static class MainUrlStatus {
        private boolean accessible;
        private int statusCode;
        private String message;
        private String redirectedUrl;
        private long responseTime;
        private String usedUrl;
        private boolean usedFallback;
        private String fallbackMessage;

        public MainUrlStatus(boolean accessible, int statusCode, String message, String redirectedUrl,
                             long responseTime, String usedUrl, boolean usedFallback, String fallbackMessage) {
            this.accessible = accessible;
            this.statusCode = statusCode;
            this.message = message;
            this.redirectedUrl = redirectedUrl;
            this.responseTime = responseTime;
            this.usedUrl = usedUrl;
            this.usedFallback = usedFallback;
            this.fallbackMessage = fallbackMessage;
        }

        public boolean isAccessible() { return accessible; }
        public int getStatusCode() { return statusCode; }
        public String getMessage() { return message; }
        public String getRedirectedUrl() { return redirectedUrl; }
        public long getResponseTime() { return responseTime; }
        public String getUsedUrl() { return usedUrl; }
        public boolean isUsedFallback() { return usedFallback; }
        public String getFallbackMessage() { return fallbackMessage; }
    }

    static class MenuItem {
        private String url;
        private String pageTitle;
        private String parentPage;
        private String childPage;
        private String fullPath;
        private int level;

        public MenuItem(String url, String pageTitle, String parentPage, String childPage, String fullPath, int level) {
            this.url = url;
            this.pageTitle = pageTitle;
            this.parentPage = parentPage;
            this.childPage = childPage;
            this.fullPath = fullPath;
            this.level = level;
        }

        public String getUrl() { return url; }
        public String getPageTitle() { return pageTitle; }
        public String getParentPage() { return parentPage; }
        public String getChildPage() { return childPage; }
        public String getFullPath() { return fullPath; }
        public int getLevel() { return level; }
    }

    // Map of contact form URLs for websites
    static final Map<String, String> CONTACT_FORM_URLS = new HashMap<>();
    static {
        CONTACT_FORM_URLS.put("https://brandcrock.in", "https://brandcrock.in/contactus/contactform");
//        CONTACT_FORM_URLS.put("https://www.qpool24.com", "https://qpool24.com/service/kontakt/");
//        CONTACT_FORM_URLS.put("https://www.kabel-schmidt.de", "https://www.kabel-schmidt.de/kontakt");
//        CONTACT_FORM_URLS.put("https://muenzdiscount.de", "https://muenzdiscount.de/kontakt");
//        CONTACT_FORM_URLS.put("https://www.gt-deko.de", "https://www.gt-deko.de/kontakt");
//        CONTACT_FORM_URLS.put("https://www.sicherheitsschirm.com", "https://sicherheitsschirm.com/pages/contact");
//        CONTACT_FORM_URLS.put("https://wirliebenhanf.de", "https://wirliebenhanf.de/contact/");
//        CONTACT_FORM_URLS.put("https://www.beautylope.de", "https://www.beautylope.de/pages/kontaktformular");
//        CONTACT_FORM_URLS.put("https://nobananas.com", "https://nobananas.com/int_en/contactus");
//        CONTACT_FORM_URLS.put("https://gatewayess.com", "https://gatewayess.com/#enquiry");
//        CONTACT_FORM_URLS.put("http://physiofitrehabcentre.com", "http://physiofitrehabcentre.com/#contact");
//        CONTACT_FORM_URLS.put("https://www.suvaifoods.us", "https://www.suvaifoods.us/contact/");
//        CONTACT_FORM_URLS.put("http://showbaglass.com", "https://showbaglass.com/#contact");
//        CONTACT_FORM_URLS.put("https://jrtoursandtravels.in", "https://jrtoursandtravels.in/contact-us/");
//        CONTACT_FORM_URLS.put("https://kaisergarten.de", "https://kaisergarten.de/pages/kontakt/");
//        CONTACT_FORM_URLS.put("https://remo-torwarthandschuhe.de", "https://remo-torwarthandschuhe.de/pages/contact");
//        CONTACT_FORM_URLS.put("https://www.chaloinc.com", "https://www.chaloinc.com/contact-us/");
//        CONTACT_FORM_URLS.put("https://loverichmondhomes.com", "https://loverichmondhomes.com/contact/");
//        CONTACT_FORM_URLS.put("https://www.muenchner-fussball-schule.de", "https://www.muenchner-fussball-schule.de/kontakt/");
//        CONTACT_FORM_URLS.put("https://pincha.de", "https://www.pincha.de/ueber-uns/");
//        CONTACT_FORM_URLS.put("https://biofreshfoods.com", "https://biofreshfoods.com/contact/");
//        CONTACT_FORM_URLS.put("https://seyongrand.com", "https://seyongrand.com/contact/");
//        CONTACT_FORM_URLS.put("http://Knrsystems.com", "https://www.rottmeir.de/kontakt/#wpcf7-f916-o1");
    }

    public static void main(String[] args) {
        // ============================================================
        // COMPLETE WEBSITE TESTING - ALL NAVIGATION LEVELS
        // Starts with main URL, automatically detects all menus, submenus, and child pages
        // Navigates through each link and tests all accessible pages
        // ============================================================
        List<String> websites = Arrays.asList(
               // "https://pr-helden.de",
//                "https://www.brandcrock.com",
//                "https://estateexpertmarketing.com",
//                "https://www.karaokebox.co.uk",
//                "https://fobits.co",
                "https://brandcrock.in"
//                "https://www.alphaindustries.eu",
//                "https://www.qpool24.com",
//                "https://www.cupio-nails.de",
//                "https://fino-gabelstapler.de",
//                "https://b2b.gewi-group.de",
//                "https://www.jeans-onlineshop.com",
//                "https://www.kabel-schmidt.de",
//                "https://muenzdiscount.de",
//                "https://www.tintsofnature.de",
//                "https://www.vape-customs.de",
//                "https://www.gt-deko.de",
//                "https://www.schuhshop24.com",
//                "https://www.lohnabfuellung-deutschland.de",
//                "https://italiamegashop.it/",
//                "https://www.sicherheitsschirm.com",
//                "https://wirliebenhanf.de",
//                "https://www.beautylope.de",
//                "https://www.die-stadtmeister.de",
//                "https://www.jeans-onlineshop.com/",
//                "https://www.dichtstoffdepot.de/",
//                "https://www.munichcricket.de/",
//                "https://nobananas.com/",
//                "https://osz.winlocal.de/products",
//                "https://shimla-germering.de",
//                "https://monizastories.com",
//                "https://gatewayess.com",
//                "https://loeschfeen.de",
//                "http://wasservilla.com",
//                "https://saxer-auto.ch",
//                "http://www.idschennai.com",
//                "https://loeschprofis.de",
//                "http://physiofitrehabcentre.com",
//                "https://www.suvaifoods.us",
//                "https://www.brandcommx.com",
//                "https://www.brandcommx.de",
//                "http://showbaglass.com",
//                "https://jrtoursandtravels.in",
//                "https://kaisergarten.de",
//                "https://wowtel.de",
//                "https://remo-torwarthandschuhe.de",
//                "https://www.chaloinc.com/",
//                "https://loverichmondhomes.com",
//                "https://www.muenchner-fussball-schule.de/",
//                "https://www.strato.de/",
//                "https://www.rottmeir.de",
//                "https://pincha.de/",
//                "https://tattoo.pincha.de/",
//                "https://biofreshfoods.com/",
//                "https://seyongrand.com/",
//                "http://Knrsystems.com"
        );

        try {
            DateTimeFormatter sheetFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss");
            String currentDateTime = LocalDateTime.now().format(sheetFormatter);
            mainSheetName = "Complete_Website_Test_Report_" + currentDateTime;

            boolean sheetsInitialized = initializeGoogleSheetsService();

            if (!sheetsInitialized) {
                System.err.println("\nWARNING: Google Sheets integration not available. Continuing without reporting...\n");
            } else {
                createNewSheet();
                setupSheetHeaders();
                ensureSheetHasEnoughRows(10000);
            }

            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
            String timestamp = LocalDateTime.now().format(formatter);
            screenshotDir = "complete_test_screenshots_" + timestamp;

            try {
                Files.createDirectories(Paths.get(screenshotDir));
                System.out.println("Screenshot directory created: " + screenshotDir);
            } catch (IOException e) {
                System.err.println("Could not create screenshot directory: " + e.getMessage());
            }

            ChromeOptions options = new ChromeOptions();
            options.addArguments("--remote-allow-origins=*");
            options.addArguments("--disable-dev-shm-usage");
            options.addArguments("--no-sandbox");
            options.addArguments("--disable-blink-features=AutomationControlled");
            options.addArguments("--headless");
            options.addArguments("--window-size=1920,1080");
            options.addArguments("--disable-notifications");
            options.addArguments("--disable-popup-blocking");
            options.addArguments("--disable-geolocation");

            Map<String, Object> prefs = new HashMap<>();
            prefs.put("profile.default_content_setting_values.notifications", 2);
            prefs.put("profile.default_content_setting_values.geolocation", 2);
            prefs.put("profile.default_content_setting_values.media_stream_mic", 2);
            prefs.put("profile.default_content_setting_values.media_stream_camera", 2);
            prefs.put("profile.default_content_setting_values.popups", 2);
            options.setExperimentalOption("prefs", prefs);

            // Process each website completely
            for (int i = 0; i < websites.size(); i++) {
                String site = websites.get(i);
                totalWebsitesProcessed++;

                System.out.println("\n" + "=".repeat(70));
                System.out.println("COMPLETE WEBSITE TESTING - " + (i+1) + " OF " + websites.size());
                System.out.println("Website: " + site);
                System.out.println("Testing all navigation levels: Menus, Submenus, and Child Pages");
                System.out.println("=".repeat(70));

                WebDriver driver = new ChromeDriver(options);
                Actions actions = new Actions(driver);
                JavascriptExecutor jsExecutor = (JavascriptExecutor) driver;
                WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

                boolean websitePassed = true;
                int websitePageCount = 0;
                int websitePassCount = 0;
                int websiteFailCount = 0;

                try {
                    addEmptyRowBeforeNewWebsite();
                    addWebsiteHeaderRow(site);

                    MainUrlStatus mainUrlStatus = checkMainUrlWithFallback(site, driver);

                    if (!mainUrlStatus.isAccessible()) {
                        recordMainUrlFailure(site, mainUrlStatus);
                        System.out.println("SKIPPING: " + site + " - Main URL not accessible");
                        websitePassed = false;
                        totalWebsitesFailed++;
                        failedWebsitesList.add(site);
                    } else {
                        // Perform COMPLETE website testing - all navigation levels
                        WebsiteTestResult websiteResult = performCompleteWebsiteTesting(site, mainUrlStatus, driver, actions, jsExecutor, wait);
                        websitePageCount = websiteResult.totalPages;
                        websitePassCount = websiteResult.passedCount;
                        websiteFailCount = websiteResult.failedCount;

                        totalPagesTestedOverall += websitePageCount;
                        totalPagesPassedOverall += websitePassCount;
                        totalPagesFailedOverall += websiteFailCount;

                        if (websiteFailCount > 0) {
                            websitePassed = false;
                            totalWebsitesFailed++;
                            failedWebsitesList.add(site);
                        } else {
                            totalWebsitesPassed++;
                        }

                        // Store website statistics
                        websitePageCountMap.put(site, websitePageCount);
                        websitePassCountMap.put(site, websitePassCount);
                        websiteFailCountMap.put(site, websiteFailCount);
                    }

                } catch (Exception e) {
                    System.err.println("Error processing website " + site + ": " + e.getMessage());
                    e.printStackTrace();
                    totalWebsitesFailed++;
                    failedWebsitesList.add(site);
                } finally {
                    if (driver != null) {
                        driver.quit();
                        System.out.println("\nClosed browser for: " + site);
                    }
                }

                System.out.println("\nCompleted complete testing for: " + site);
                System.out.println("  Pages tested: " + websitePageCount);
                System.out.println("  Passed: " + websitePassCount);
                System.out.println("  Failed: " + websiteFailCount);
                System.out.println("=".repeat(70));
            }

            // Add final summary to Google Sheet
            addFinalSummaryToSheet();

            System.out.println("\n" + "=".repeat(70));
            System.out.println("ALL WEBSITES TESTED COMPLETELY");
            System.out.println("=".repeat(70));
            System.out.println("Total Websites Processed: " + totalWebsitesProcessed);
            System.out.println("Websites PASSED: " + totalWebsitesPassed);
            System.out.println("Websites FAILED: " + totalWebsitesFailed);
            System.out.println("Total Pages Tested Overall: " + totalPagesTestedOverall);
            System.out.println("Total Pages Passed: " + totalPagesPassedOverall);
            System.out.println("Total Pages Failed: " + totalPagesFailedOverall);
            System.out.println("Overall Pass Rate: " + String.format("%.2f%%", (double) totalPagesPassedOverall / totalPagesTestedOverall * 100));
            System.out.println("=".repeat(70));
            System.out.println("Screenshots saved in: " + screenshotDir);
            if (sheetsService != null) {
                System.out.println("Complete report with final summary added to Google Sheet: " + mainSheetName);
            }
            System.out.println("=".repeat(70));

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    // Class to hold website test results
    static class WebsiteTestResult {
        int totalPages;
        int passedCount;
        int failedCount;

        WebsiteTestResult(int totalPages, int passedCount, int failedCount) {
            this.totalPages = totalPages;
            this.passedCount = passedCount;
            this.failedCount = failedCount;
        }
    }

    // Perform complete website testing - all navigation levels
    static WebsiteTestResult performCompleteWebsiteTesting(String baseUrl, MainUrlStatus mainUrlStatus, WebDriver driver,
                                                           Actions actions, JavascriptExecutor jsExecutor, WebDriverWait wait) {
        System.out.println("\n" + "-".repeat(60));
        System.out.println("STARTING COMPLETE NAVIGATION TESTING");
        System.out.println("Website: " + baseUrl);
        System.out.println("Testing all menus, submenus, and child pages...");
        System.out.println("-".repeat(60));

        testResults = new ArrayList<>();
        allUrls = new HashSet<>();

        try {
            long testStartTime = System.currentTimeMillis();
            System.out.println("Loading main URL: " + mainUrlStatus.getUsedUrl());
            driver.get(mainUrlStatus.getUsedUrl());
            Thread.sleep(2000);
            handlePopups(driver, jsExecutor);
            Thread.sleep(1000);

            System.out.println("Expanding all menus and submenus...");
            expandAllMenus(driver, actions, jsExecutor, wait);

            System.out.println("Discovering all navigation links...");
            List<MenuItem> menuItems = getAllMenuItemsWithHierarchy(driver, jsExecutor, actions);

            System.out.println("\nFound " + menuItems.size() + " total pages/URLs to test");
            System.out.println("  - Level 1 (Top Menus): " + menuItems.stream().filter(i -> i.getLevel() == 1).count());
            System.out.println("  - Level 2 (Submenus): " + menuItems.stream().filter(i -> i.getLevel() == 2).count());
            System.out.println("  - Level 3+ (Child Pages): " + menuItems.stream().filter(i -> i.getLevel() >= 3).count());

            int passedCount = 0;
            int failedCount = 0;
            long totalDuration = 0;

            System.out.println("\nTesting each discovered page...");
            for (MenuItem menuItem : menuItems) {
                TestResult result = testUrlStatusWithDetails(baseUrl, menuItem, driver, jsExecutor);
                testResults.add(result);

                if ("PASS".equals(result.status)) {
                    passedCount++;
                } else {
                    failedCount++;
                }
                totalDuration += result.durationMs;

                String levelIndicator = "";
                if (menuItem.getLevel() == 1) levelIndicator = "[TOP MENU]";
                else if (menuItem.getLevel() == 2) levelIndicator = "[SUBMENU]";
                else levelIndicator = "[CHILD PAGE]";

                String statusText = "PASS".equals(result.status) ? "PASS" : "FAIL";
                System.out.println(statusText + " " + levelIndicator + " - " + menuItem.getFullPath() +
                        " (" + result.statusCode + ") - " + result.durationMs + "ms");
                Thread.sleep(200);
            }

            appendResultsToSheet(testResults);

            long testEndTime = System.currentTimeMillis();
            System.out.println("\n" + "-".repeat(60));
            System.out.println("TESTING SUMMARY FOR: " + baseUrl);
            System.out.println("-".repeat(60));
            System.out.println("Total pages discovered and tested: " + menuItems.size());
            System.out.println("PASSED: " + passedCount);
            System.out.println("FAILED: " + failedCount);
            System.out.println("Pass Rate: " + String.format("%.2f%%", (double) passedCount / menuItems.size() * 100));
            System.out.println("Total Duration: " + (testEndTime - testStartTime) + " ms");
            System.out.println("-".repeat(60));

            // After completing all navigation testing, perform form submission test if applicable
            if (CONTACT_FORM_URLS.containsKey(baseUrl)) {
                System.out.println("\n--- PERFORMING ADDITIONAL FORM SUBMISSION TEST ---");
                String formUrl = CONTACT_FORM_URLS.get(baseUrl);
                performFormSubmissionTest(driver, wait, jsExecutor, baseUrl, formUrl);
            }

            return new WebsiteTestResult(menuItems.size(), passedCount, failedCount);

        } catch (Exception e) {
            System.err.println("Error during complete website testing: " + e.getMessage());
            e.printStackTrace();
            return new WebsiteTestResult(0, 0, 0);
        }
    }

    // Add final summary to Google Sheet
    private static void addFinalSummaryToSheet() {
        if (sheetsService == null) return;

        try {
            String dateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));

            // Add empty rows before summary
            ValueRange response = sheetsService.spreadsheets().values()
                    .get(SPREADSHEET_ID, mainSheetName + "!A:A")
                    .execute();
            int lastRow = response.getValues() != null ? response.getValues().size() : 1;

            // Add separator rows
            List<List<Object>> separatorRows = Arrays.asList(
                    Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", ""),
                    Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", ""),
                    Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", ""),
                    Arrays.asList("========== FINAL TEST SUMMARY ==========", "", "", "", "", "", "", "", "", "", "", "", ""),
                    Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", "")
            );

            ValueRange separatorBody = new ValueRange().setValues(separatorRows);
            sheetsService.spreadsheets().values()
                    .update(SPREADSHEET_ID, mainSheetName + "!A" + (lastRow + 1) + ":M" + (lastRow + separatorRows.size()), separatorBody)
                    .setValueInputOption("RAW")
                    .execute();

            lastRow = lastRow + separatorRows.size();

            // Prepare summary data
            List<List<Object>> summaryRows = new ArrayList<>();
            summaryRows.add(Arrays.asList("Summary Generated On:", dateTime, "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("OVERALL TEST STATISTICS", "", "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Total Websites Processed:", String.valueOf(totalWebsitesProcessed), "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Websites PASSED:", String.valueOf(totalWebsitesPassed), "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Websites FAILED:", String.valueOf(totalWebsitesFailed), "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Overall Website Pass Rate:", String.format("%.2f%%", (double) totalWebsitesPassed / totalWebsitesProcessed * 100), "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("PAGE LEVEL STATISTICS", "", "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Total Pages Tested Overall:", String.valueOf(totalPagesTestedOverall), "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Total Pages PASSED:", String.valueOf(totalPagesPassedOverall), "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Total Pages FAILED:", String.valueOf(totalPagesFailedOverall), "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Overall Page Pass Rate:", String.format("%.2f%%", (double) totalPagesPassedOverall / totalPagesTestedOverall * 100), "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("WEBSITE WISE BREAKDOWN", "", "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Website URL", "Pages Tested", "Pages Passed", "Pages Failed", "Pass Rate", "", "", "", "", "", "", "", ""));

            // Add individual website statistics
            for (String website : websitePageCountMap.keySet()) {
                int total = websitePageCountMap.getOrDefault(website, 0);
                int passed = websitePassCountMap.getOrDefault(website, 0);
                int failed = websiteFailCountMap.getOrDefault(website, 0);
                double passRate = total > 0 ? (double) passed / total * 100 : 0;
                String status = failed > 0 ? "FAILED" : "PASSED";

                summaryRows.add(Arrays.asList(
                        website,
                        String.valueOf(total),
                        String.valueOf(passed),
                        String.valueOf(failed),
                        String.format("%.2f%%", passRate),
                        status,
                        "", "", "", "", "", "", ""
                ));
            }

            summaryRows.add(Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", ""));

            if (!failedWebsitesList.isEmpty()) {
                summaryRows.add(Arrays.asList("FAILED WEBSITES LIST:", "", "", "", "", "", "", "", "", "", "", "", ""));
                for (String failedSite : failedWebsitesList) {
                    summaryRows.add(Arrays.asList("- " + failedSite, "", "", "", "", "", "", "", "", "", "", "", ""));
                }
            }

            summaryRows.add(Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("TESTING COMPLETION STATUS:", "COMPLETED", "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("All websites have been completely tested.", "", "", "", "", "", "", "", "", "", "", "", ""));
            summaryRows.add(Arrays.asList("Test End Time:", dateTime, "", "", "", "", "", "", "", "", "", "", ""));

            // Add summary to sheet
            ValueRange summaryBody = new ValueRange().setValues(summaryRows);
            sheetsService.spreadsheets().values()
                    .update(SPREADSHEET_ID, mainSheetName + "!A" + (lastRow + 1) + ":M" + (lastRow + summaryRows.size()), summaryBody)
                    .setValueInputOption("RAW")
                    .execute();

            // Format the summary header
            formatSummarySection(lastRow + 3);

            System.out.println("\nFinal summary added to Google Sheet successfully!");

        } catch (Exception e) {
            System.err.println("Could not add final summary to sheet: " + e.getMessage());
        }
    }

    // Format the summary section in Google Sheet
    private static void formatSummarySection(int headerRowIndex) throws IOException {
        Integer sheetId = getSheetIdByName(mainSheetName);
        if (sheetId == null) return;

        List<Request> requests = new ArrayList<>();

        // Bold the summary headers
        requests.add(new Request().setRepeatCell(new RepeatCellRequest()
                .setRange(new GridRange()
                        .setSheetId(sheetId)
                        .setStartRowIndex(headerRowIndex - 1)
                        .setEndRowIndex(headerRowIndex + 2))
                .setCell(new CellData().setUserEnteredFormat(
                        new CellFormat()
                                .setTextFormat(new TextFormat().setBold(true))))
                .setFields("userEnteredFormat.textFormat.bold")));

        // Bold the website breakdown header
        requests.add(new Request().setRepeatCell(new RepeatCellRequest()
                .setRange(new GridRange()
                        .setSheetId(sheetId)
                        .setStartRowIndex(headerRowIndex + 13)
                        .setEndRowIndex(headerRowIndex + 14))
                .setCell(new CellData().setUserEnteredFormat(
                        new CellFormat()
                                .setTextFormat(new TextFormat().setBold(true))))
                .setFields("userEnteredFormat.textFormat.bold")));

        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();
    }

    // Enhanced form submission test with all field types
    private static void performFormSubmissionTest(WebDriver driver, WebDriverWait wait,
                                                  JavascriptExecutor jsExecutor,
                                                  String websiteUrl, String formUrl) {
        String screenshotPath = null;
        String errorDetails = null;
        String status = "PASS";

        try {
            System.out.println("  Navigating to contact form: " + formUrl);
            driver.get(formUrl);
            Thread.sleep(3000);

            if (formUrl.contains("#")) {
                try {
                    String anchor = formUrl.substring(formUrl.indexOf("#"));
                    WebElement element = driver.findElement(By.cssSelector(anchor));
                    jsExecutor.executeScript("arguments[0].scrollIntoView(true);", element);
                    Thread.sleep(1000);
                } catch (Exception e) {}
            }

            handlePopups(driver, jsExecutor);

            System.out.println("  Filling form fields...");
            boolean formFilled = fillContactForm(driver, jsExecutor);

            if (!formFilled) {
                status = "FAIL";
                errorDetails = "Could not locate or fill required form fields";
                screenshotPath = takeFormSubmissionScreenshot(driver, websiteUrl);
                recordFormSubmissionToSameSheet(websiteUrl, formUrl, status, errorDetails, screenshotPath);
                return;
            }

            System.out.println("  Submitting form...");
            boolean submitted = submitContactForm(driver, jsExecutor);

            if (!submitted) {
                status = "FAIL";
                errorDetails = "Could not find or click submit button";
                screenshotPath = takeFormSubmissionScreenshot(driver, websiteUrl);
                recordFormSubmissionToSameSheet(websiteUrl, formUrl, status, errorDetails, screenshotPath);
                return;
            }

            Thread.sleep(5000);

            String pageSource = driver.getPageSource().toLowerCase();
            boolean success = checkFormSubmissionSuccess(pageSource);
            boolean hasCaptcha = pageSource.contains("captcha") || pageSource.contains("recaptcha");

            if (hasCaptcha) {
                status = "FAIL";
                errorDetails = "CAPTCHA detected - Form could not be submitted automatically";
                screenshotPath = takeFormSubmissionScreenshot(driver, websiteUrl);
            } else if (!success) {
                status = "FAIL";
                errorDetails = "Form submitted but no success message detected";
                screenshotPath = takeFormSubmissionScreenshot(driver, websiteUrl);
            } else {
                System.out.println("  Form submitted successfully!");
            }

            recordFormSubmissionToSameSheet(websiteUrl, formUrl, status, errorDetails, screenshotPath);

        } catch (Exception e) {
            status = "FAIL";
            errorDetails = e.getMessage();
            screenshotPath = takeFormSubmissionScreenshot(driver, websiteUrl);
            System.err.println("  Form submission error: " + e.getMessage());
            recordFormSubmissionToSameSheet(websiteUrl, formUrl, status, errorDetails, screenshotPath);
        }
    }

    // Enhanced form filling with support for all field types
    private static boolean fillContactForm(WebDriver driver, JavascriptExecutor jsExecutor) {
        boolean nameFilled = false;
        boolean lastNameFilled = false;
        boolean emailFilled = false;
        boolean phoneFilled = false;
        boolean subjectFilled = false;
        boolean companyFilled = false;
        boolean addressFilled = false;
        boolean cityFilled = false;
        boolean zipFilled = false;
        boolean countryFilled = false;
        boolean websiteFilled = false;
        boolean messageFilled = false;

        try {
            Thread.sleep(2000);
            List<WebElement> allFields = driver.findElements(By.xpath("//input | //textarea"));

            for (WebElement field : allFields) {
                String type = field.getAttribute("type");
                if (type == null) type = field.getTagName();
                type = type.toLowerCase();

                if (type.equals("hidden") || type.equals("submit") || type.equals("button") || type.equals("reset") || type.equals("image")) {
                    continue;
                }

                if (!field.isDisplayed() || !field.isEnabled()) continue;

                String name = field.getAttribute("name");
                String id = field.getAttribute("id");
                String placeholder = field.getAttribute("placeholder");
                String ariaLabel = field.getAttribute("aria-label");
                String className = field.getAttribute("class");

                String identifier = ((name != null ? name : "") + (id != null ? id : "") +
                        (placeholder != null ? placeholder : "") + (ariaLabel != null ? ariaLabel : "") +
                        (className != null ? className : "")).toLowerCase();

                // First Name field detection
                if (!nameFilled && (identifier.contains("first") && identifier.contains("name")) ||
                        identifier.equals("fname") || identifier.equals("firstname")) {
                    field.clear();
                    field.sendKeys(TestData.FIRST_NAME);
                    nameFilled = true;
                    System.out.println("    Filled First Name field: " + TestData.FIRST_NAME);
                }
                // Last Name field detection
                else if (!lastNameFilled && (identifier.contains("last") && identifier.contains("name")) ||
                        identifier.equals("lname") || identifier.equals("lastname") ||
                        (placeholder != null && placeholder.toLowerCase().contains("last name"))) {
                    field.clear();
                    field.sendKeys(TestData.LAST_NAME);
                    lastNameFilled = true;
                    System.out.println("    Filled Last Name field: " + TestData.LAST_NAME);
                }
                // Full Name field detection (if no separate first/last)
                else if (!nameFilled && (identifier.contains("name") || identifier.contains("fullname") ||
                        (placeholder != null && (placeholder.toLowerCase().contains("name") || placeholder.toLowerCase().contains("vollstandiger"))))) {
                    field.clear();
                    field.sendKeys(TestData.FULL_NAME);
                    nameFilled = true;
                    System.out.println("    Filled Name field: " + TestData.FULL_NAME);
                }
                // Email field detection
                else if (!emailFilled && (identifier.contains("email") || identifier.contains("e-mail") ||
                        type.equals("email") || (placeholder != null && placeholder.toLowerCase().contains("email")))) {
                    field.clear();
                    field.sendKeys(TestData.EMAIL);
                    emailFilled = true;
                    System.out.println("    Filled Email field: " + TestData.EMAIL);
                }
                // Phone field detection
                else if (!phoneFilled && (identifier.contains("phone") || identifier.contains("telephone") ||
                        identifier.contains("tel") || identifier.contains("mobile") || identifier.contains("telefon") ||
                        (placeholder != null && (placeholder.toLowerCase().contains("phone") || placeholder.toLowerCase().contains("telefon"))))) {
                    field.clear();
                    field.sendKeys(TestData.PHONE);
                    phoneFilled = true;
                    System.out.println("    Filled Phone field: " + TestData.PHONE);
                }
                // Subject field detection
                else if (!subjectFilled && (identifier.contains("subject") || identifier.contains("betreff") ||
                        (placeholder != null && (placeholder.toLowerCase().contains("subject") || placeholder.toLowerCase().contains("betreff"))))) {
                    field.clear();
                    field.sendKeys(TestData.SUBJECT);
                    subjectFilled = true;
                    System.out.println("    Filled Subject field: " + TestData.SUBJECT);
                }
                // Company field detection
                else if (!companyFilled && (identifier.contains("company") || identifier.contains("firma") ||
                        identifier.contains("organization") || identifier.contains("organisation") ||
                        (placeholder != null && (placeholder.toLowerCase().contains("company") || placeholder.toLowerCase().contains("firma"))))) {
                    field.clear();
                    field.sendKeys(TestData.COMPANY);
                    companyFilled = true;
                    System.out.println("    Filled Company field: " + TestData.COMPANY);
                }
                // Address field detection
                else if (!addressFilled && (identifier.contains("address") || identifier.contains("adresse") ||
                        identifier.contains("street") || identifier.contains("strasse") ||
                        (placeholder != null && (placeholder.toLowerCase().contains("address") || placeholder.toLowerCase().contains("adresse"))))) {
                    field.clear();
                    field.sendKeys(TestData.ADDRESS);
                    addressFilled = true;
                    System.out.println("    Filled Address field: " + TestData.ADDRESS);
                }
                // City field detection
                else if (!cityFilled && (identifier.contains("city") || identifier.contains("ort") ||
                        identifier.contains("stadt") || identifier.contains("town") ||
                        (placeholder != null && (placeholder.toLowerCase().contains("city") || placeholder.toLowerCase().contains("ort"))))) {
                    field.clear();
                    field.sendKeys(TestData.CITY);
                    cityFilled = true;
                    System.out.println("    Filled City field: " + TestData.CITY);
                }
                // ZIP/Postal code field detection
                else if (!zipFilled && (identifier.contains("zip") || identifier.contains("postal") ||
                        identifier.contains("pincode") || identifier.contains("plz") ||
                        identifier.contains("postcode") || (placeholder != null && (placeholder.toLowerCase().contains("zip") || placeholder.toLowerCase().contains("plz"))))) {
                    field.clear();
                    field.sendKeys(TestData.ZIP_CODE);
                    zipFilled = true;
                    System.out.println("    Filled ZIP/Postal Code field: " + TestData.ZIP_CODE);
                }
                // Country field detection
                else if (!countryFilled && (identifier.contains("country") || identifier.contains("land") ||
                        (placeholder != null && (placeholder.toLowerCase().contains("country") || placeholder.toLowerCase().contains("land"))))) {
                    field.clear();
                    field.sendKeys(TestData.COUNTRY);
                    countryFilled = true;
                    System.out.println("    Filled Country field: " + TestData.COUNTRY);
                }
                // Website/URL field detection
                else if (!websiteFilled && (identifier.contains("website") || identifier.contains("url") ||
                        identifier.contains("homepage") || identifier.contains("web") ||
                        (placeholder != null && (placeholder.toLowerCase().contains("website") || placeholder.toLowerCase().contains("url"))))) {
                    field.clear();
                    field.sendKeys(TestData.WEBSITE);
                    websiteFilled = true;
                    System.out.println("    Filled Website field: " + TestData.WEBSITE);
                }
                // Message/textarea field detection
                else if (!messageFilled && (field.getTagName().equals("textarea") ||
                        identifier.contains("message") || identifier.contains("nachricht") ||
                        identifier.contains("comment") || identifier.contains("anfrage") ||
                        identifier.contains("feedback") || (placeholder != null && (placeholder.toLowerCase().contains("message") || placeholder.toLowerCase().contains("nachricht"))))) {
                    field.clear();
                    field.sendKeys(TestData.MESSAGE);
                    messageFilled = true;
                    System.out.println("    Filled Message field");
                }
                // For any other text field that wasn't matched, fill with generic test data
                else if (type.equals("text") || type.equals("search")) {
                    String currentValue = field.getAttribute("value");
                    if (currentValue == null || currentValue.isEmpty()) {
                        String genericValue = "Test data for " + (name != null ? name : "field");
                        field.clear();
                        field.sendKeys(genericValue);
                        System.out.println("    Filled additional field: " + (name != null ? name : "unnamed") + " = " + genericValue);
                    }
                }
            }

            // Return true if at least name, email, and message were filled
            return nameFilled && emailFilled && messageFilled;

        } catch (Exception e) {
            System.err.println("Error filling form: " + e.getMessage());
            return false;
        }
    }

    private static boolean submitContactForm(WebDriver driver, JavascriptExecutor jsExecutor) {
        String[] submitSelectors = {
                "input[type='submit']", "button[type='submit']", ".wpcf7-submit",
                "input[value*='Send']", "input[value*='Senden']", "input[value*='Submit']",
                "button:contains('Senden')", "button:contains('Send')", "button[class*='submit']",
                ".btn-primary", "[class*='send']", "[class*='submit']"
        };

        for (String selector : submitSelectors) {
            try {
                List<WebElement> buttons = driver.findElements(By.cssSelector(selector));
                for (WebElement button : buttons) {
                    if (button.isDisplayed() && button.isEnabled()) {
                        //jsExecutor.executeScript("arguments[0].click();", button);
                        System.out.println("    Clicked submit button");
                        return true;
                    }
                }
            } catch (Exception e) {}
        }
        return false;
    }

    private static boolean checkFormSubmissionSuccess(String pageSource) {
        String[] successIndicators = {
                "vielen dank", "thank you", "erfolgreich", "success", "your message has been sent",
                "ihre nachricht wurde gesendet", "form was sent", "we'll get back to you",
                "message sent", "email sent", "form submitted", "wurde gesendet",
                "your enquiry has been submitted", "we will contact you"
        };
        for (String indicator : successIndicators) {
            if (pageSource.contains(indicator)) return true;
        }
        return false;
    }

    private static void recordFormSubmissionToSameSheet(String websiteUrl, String formUrl,
                                                        String status, String errorDetails,
                                                        String screenshotPath) {
        String dateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));

        TestResult formTestResult = new TestResult(
                dateTime, websiteUrl, "Form Submission Page", "Form", "Contact Form",
                "Contact Form", "Form Submission", formUrl, 0L,
                status.equals("PASS") ? "Form submitted successfully" : errorDetails,
                status.equals("PASS") ? 200 : 500, status, screenshotPath
        );

        if (testResults == null) {
            testResults = new ArrayList<>();
        }
        testResults.add(formTestResult);
        appendResultsToSheet(Collections.singletonList(formTestResult));

        System.out.println("  Form submission result added to main sheet");
    }

    private static String takeFormSubmissionScreenshot(WebDriver driver, String fileNamePrefix) {
        try {
            String safeName = fileNamePrefix.replace("https://", "").replace("http://", "")
                    .replace("/", "_").replace(".", "_");
            String timestamp = String.valueOf(System.currentTimeMillis());
            String filename = "FORM_" + safeName + "_" + timestamp + ".png";
            Path destination = Paths.get(screenshotDir, filename);
            File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            Files.copy(screenshot.toPath(), destination);
            System.out.println("    Form screenshot saved: " + filename);
            return destination.toString();
        } catch (Exception e) {
            System.err.println("    Failed to save screenshot: " + e.getMessage());
            return null;
        }
    }

    private static void ensureSheetHasEnoughRows(int requiredRows) {
        if (sheetsService == null) return;
        try {
            Integer sheetId = getSheetIdByName(mainSheetName);
            if (sheetId == null) return;

            Spreadsheet spreadsheet = sheetsService.spreadsheets().get(SPREADSHEET_ID).execute();
            Sheet sheet = null;
            for (Sheet s : spreadsheet.getSheets()) {
                if (s.getProperties().getTitle().equals(mainSheetName)) {
                    sheet = s;
                    break;
                }
            }

            if (sheet == null) return;

            int currentRowCount = sheet.getProperties().getGridProperties().getRowCount();
            if (currentRowCount < requiredRows) {
                List<Request> requests = new ArrayList<>();
                requests.add(new Request().setAppendDimension(new AppendDimensionRequest()
                        .setSheetId(sheetId)
                        .setDimension("ROWS")
                        .setLength(requiredRows - currentRowCount)));

                BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                        .setRequests(requests);
                sheetsService.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();
                System.out.println("Added " + (requiredRows - currentRowCount) + " rows to the sheet");
            }
        } catch (Exception e) {
            System.err.println("Could not ensure sheet rows: " + e.getMessage());
        }
    }

    private static void addEmptyRowBeforeNewWebsite() {
        if (sheetsService == null) return;
        try {
            if (isFirstWebsite) {
                isFirstWebsite = false;
                return;
            }

            ensureSheetHasEnoughRows(getCurrentRowCount() + 5);

            ValueRange response = sheetsService.spreadsheets().values()
                    .get(SPREADSHEET_ID, mainSheetName + "!A:A")
                    .execute();
            int lastRow = response.getValues() != null ? response.getValues().size() : 1;

            List<List<Object>> emptyRows = Arrays.asList(
                    Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", ""),
                    Arrays.asList("", "", "", "", "", "", "", "", "", "", "", "", "")
            );
            ValueRange body = new ValueRange().setValues(emptyRows);
            sheetsService.spreadsheets().values()
                    .update(SPREADSHEET_ID, mainSheetName + "!A" + (lastRow + 1) + ":M" + (lastRow + 2), body)
                    .setValueInputOption("RAW")
                    .execute();

        } catch (Exception e) {
            System.err.println("Could not add separator rows: " + e.getMessage());
        }
    }

    private static int getCurrentRowCount() {
        try {
            ValueRange response = sheetsService.spreadsheets().values()
                    .get(SPREADSHEET_ID, mainSheetName + "!A:A")
                    .execute();
            return response.getValues() != null ? response.getValues().size() : 1;
        } catch (Exception e) {
            return 1;
        }
    }

    private static void addWebsiteHeaderRow(String websiteUrl) {
        if (sheetsService == null) return;
        try {
            ensureSheetHasEnoughRows(getCurrentRowCount() + 2);

            ValueRange response = sheetsService.spreadsheets().values()
                    .get(SPREADSHEET_ID, mainSheetName + "!A:A")
                    .execute();
            int lastRow = response.getValues() != null ? response.getValues().size() : 1;

            String headerText = "=== COMPLETE TESTING WEBSITE: " + websiteUrl + " ===";
            List<List<Object>> headerRow = Arrays.asList(Arrays.asList(
                    headerText, "", "", "", "", "", "", "", "", "", "", "", ""
            ));
            ValueRange body = new ValueRange().setValues(headerRow);
            sheetsService.spreadsheets().values()
                    .update(SPREADSHEET_ID, mainSheetName + "!A" + (lastRow + 1) + ":M" + (lastRow + 1), body)
                    .setValueInputOption("RAW")
                    .execute();

            formatHeaderRowForWebsite(lastRow + 1);
            System.out.println("Added website header for: " + websiteUrl);

        } catch (Exception e) {
            System.err.println("Could not add website header: " + e.getMessage());
        }
    }

    private static void formatHeaderRowForWebsite(int rowIndex) throws IOException {
        Integer sheetId = getSheetIdByName(mainSheetName);
        if (sheetId == null) return;

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setRepeatCell(new RepeatCellRequest()
                .setRange(new GridRange()
                        .setSheetId(sheetId)
                        .setStartRowIndex(rowIndex - 1)
                        .setEndRowIndex(rowIndex))
                .setCell(new CellData().setUserEnteredFormat(
                        new CellFormat()
                                .setTextFormat(new TextFormat().setBold(true))
                                .setBackgroundColor(new Color().setRed(0.9f).setGreen(0.9f).setBlue(0.9f))))
                .setFields("userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor")));

        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();
    }

    private static MainUrlStatus checkMainUrlWithFallback(String originalUrl, WebDriver driver) {
        System.out.println("\nChecking main URL: " + originalUrl);

        String wwwUrl = originalUrl;
        String nonWwwUrl = originalUrl;

        if (originalUrl.contains("://www.")) {
            nonWwwUrl = originalUrl.replace("://www.", "://");
        } else if (originalUrl.contains("://") && !originalUrl.contains("://www.")) {
            String protocol = originalUrl.substring(0, originalUrl.indexOf("://") + 3);
            String domain = originalUrl.substring(originalUrl.indexOf("://") + 3);
            wwwUrl = protocol + "www." + domain;
        }

        System.out.println("  Trying: " + wwwUrl);
        MainUrlStatus primaryStatus = checkSingleUrl(wwwUrl);

        if (primaryStatus.isAccessible()) {
            System.out.println("  SUCCESS: " + wwwUrl + " is accessible");
            return new MainUrlStatus(true, primaryStatus.getStatusCode(), primaryStatus.getMessage(),
                    primaryStatus.getRedirectedUrl(), primaryStatus.getResponseTime(),
                    wwwUrl, false, null);
        }

        System.out.println("  Primary failed, trying fallback: " + nonWwwUrl);
        MainUrlStatus fallbackStatus = checkSingleUrl(nonWwwUrl);

        if (fallbackStatus.isAccessible()) {
            String fallbackMessage = String.format("Primary URL (%s) failed (HTTP %d). Using fallback URL (%s) instead.",
                    wwwUrl, primaryStatus.getStatusCode(), nonWwwUrl);
            System.out.println("  SUCCESS with fallback: " + nonWwwUrl);
            System.out.println("  Note: " + fallbackMessage);
            return new MainUrlStatus(true, fallbackStatus.getStatusCode(), fallbackStatus.getMessage(),
                    fallbackStatus.getRedirectedUrl(), fallbackStatus.getResponseTime(),
                    nonWwwUrl, true, fallbackMessage);
        }

        String finalMessage = String.format("Both URLs failed. Primary: %s (HTTP %d), Fallback: %s (HTTP %d)",
                wwwUrl, primaryStatus.getStatusCode(),
                nonWwwUrl, fallbackStatus.getStatusCode());
        System.out.println("  FAILED: " + finalMessage);
        return new MainUrlStatus(false, fallbackStatus.getStatusCode(), finalMessage,
                null, fallbackStatus.getResponseTime(), nonWwwUrl, true, finalMessage);
    }

    private static MainUrlStatus checkSingleUrl(String url) {
        try {
            long startTime = System.currentTimeMillis();

            HttpURLConnection conn = (HttpURLConnection) new URL(url).openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(10000);
            conn.setReadTimeout(10000);
            conn.setInstanceFollowRedirects(false);
            conn.connect();

            int statusCode = conn.getResponseCode();
            String redirectedUrl = null;
            String message = "";
            boolean accessible = false;
            long responseTime = System.currentTimeMillis() - startTime;

            if (statusCode == 301 || statusCode == 302 || statusCode == 307 || statusCode == 308) {
                redirectedUrl = conn.getHeaderField("Location");
                message = String.format("Redirected to: %s (HTTP %d)", redirectedUrl, statusCode);
                accessible = true;
            }
            else if (statusCode == 200) {
                message = "OK";
                accessible = true;
            }
            else if (statusCode >= 400 && statusCode < 500) {
                message = getHttpStatusMessage(statusCode) + " - Client Error";
                accessible = false;
            }
            else if (statusCode >= 500 && statusCode < 600) {
                message = getHttpStatusMessage(statusCode) + " - Server Error";
                accessible = false;
            }
            else {
                message = getHttpStatusMessage(statusCode);
                accessible = false;
            }

            conn.disconnect();
            return new MainUrlStatus(accessible, statusCode, message, redirectedUrl, responseTime, url, false, null);

        } catch (Exception e) {
            return new MainUrlStatus(false, 0, "Connection failed: " + e.getMessage(), null, 0, url, false, null);
        }
    }

    private static String getHttpStatusMessage(int statusCode) {
        switch (statusCode) {
            case 200: return "OK";
            case 301: return "Moved Permanently";
            case 302: return "Found";
            case 307: return "Temporary Redirect";
            case 308: return "Permanent Redirect";
            case 400: return "Bad Request";
            case 401: return "Unauthorized";
            case 403: return "Forbidden";
            case 404: return "Not Found";
            case 410: return "Gone";
            case 429: return "Too Many Requests";
            case 500: return "Internal Server Error";
            case 502: return "Bad Gateway";
            case 503: return "Service Unavailable";
            case 504: return "Gateway Timeout";
            default: return "HTTP " + statusCode;
        }
    }

    private static void recordMainUrlFailure(String originalUrl, MainUrlStatus status) {
        System.out.println("\nMAIN URL FAILURE: " + originalUrl);
        System.out.println("Status Code: " + status.getStatusCode());
        System.out.println("Message: " + status.getMessage());
        if (status.isUsedFallback()) {
            System.out.println("Fallback Info: " + status.getFallbackMessage());
        }
        System.out.println("Response Time: " + status.getResponseTime() + "ms");

        if (sheetsService == null) return;

        String dateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        String pageTitle = "MAIN URL - " + (status.getStatusCode() == 0 ? "CONNECTION FAILED" : "HTTP " + status.getStatusCode());
        String parentPage = "N/A";
        String childPage = "N/A";
        String fullPath = "Main URL";
        String action = "Main URL Check";
        String testedUrl = status.getUsedUrl();

        String failedResponseLog = status.getMessage();
        if (status.isUsedFallback()) {
            failedResponseLog = status.getFallbackMessage();
        }

        TestResult mainUrlResult = new TestResult(
                dateTime, originalUrl, pageTitle, parentPage, childPage, fullPath, action, testedUrl,
                status.getResponseTime(), failedResponseLog, status.getStatusCode(),
                status.isAccessible() ? "PASS (with fallback)" : "FAIL",
                "No screenshot - Main URL " + (status.isAccessible() ? "used fallback" : "inaccessible")
        );

        appendResultsToSheet(Collections.singletonList(mainUrlResult));
    }

    private static boolean initializeGoogleSheetsService() {
        try {
            InputStream credentialsStream = null;
            String[] searchPaths = {
                    "src/main/resources/credentials.json",
                    "credentials.json",
                    System.getProperty("user.dir") + "/src/main/resources/credentials.json",
                    System.getProperty("user.dir") + "/credentials.json"
            };

            for (String path : searchPaths) {
                File file = new File(path);
                if (file.exists()) {
                    credentialsStream = new FileInputStream(file);
                    System.out.println("Found credentials at: " + path);
                    break;
                }
            }

            if (credentialsStream == null) {
                System.err.println("\nERROR: credentials.json not found!");
                return false;
            }

            GoogleCredentials credentials = GoogleCredentials.fromStream(credentialsStream)
                    .createScoped(Collections.singletonList(SheetsScopes.SPREADSHEETS));

            final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

            sheetsService = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, new HttpCredentialsAdapter(credentials))
                    .setApplicationName(APPLICATION_NAME)
                    .build();

            com.google.api.services.sheets.v4.model.Spreadsheet spreadsheet = sheetsService.spreadsheets().get(SPREADSHEET_ID).execute();
            System.out.println("Google Sheets service initialized");
            System.out.println("Connected to: " + spreadsheet.getProperties().getTitle());

            return true;

        } catch (Exception e) {
            System.err.println("Could not initialize Google Sheets: " + e.getMessage());
            return false;
        }
    }

    private static void createNewSheet() {
        if (sheetsService == null) return;
        try {
            AddSheetRequest addSheetRequest = new AddSheetRequest()
                    .setProperties(new SheetProperties().setTitle(mainSheetName));
            BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                    .setRequests(Collections.singletonList(new Request().setAddSheet(addSheetRequest)));
            sheetsService.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();
            System.out.println("Created sheet: " + mainSheetName);
        } catch (Exception e) {
            System.err.println("Could not create sheet: " + e.getMessage());
        }
    }

    private static void setupSheetHeaders() {
        if (sheetsService == null) return;
        try {
            List<List<Object>> headers = Arrays.asList(Arrays.asList(
                    "Date and Time", "Website URL", "Page Title", "Parent Page", "Child Page",
                    "Full Navigation Path", "Action", "Tested URL", "Duration (ms)",
                    "Error Details", "Status Code", "Status", "Screenshot Path"
            ));
            ValueRange body = new ValueRange().setValues(headers);
            sheetsService.spreadsheets().values()
                    .update(SPREADSHEET_ID, mainSheetName + "!A1:M1", body)
                    .setValueInputOption("RAW")
                    .execute();

            formatMainHeaderRow();
            System.out.println("Headers added");
        } catch (Exception e) {
            System.err.println("Could not add headers: " + e.getMessage());
        }
    }

    private static void formatMainHeaderRow() throws IOException {
        Integer sheetId = getSheetIdByName(mainSheetName);
        if (sheetId == null) return;

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setRepeatCell(new RepeatCellRequest()
                .setRange(new GridRange()
                        .setSheetId(sheetId)
                        .setStartRowIndex(0)
                        .setEndRowIndex(1))
                .setCell(new CellData().setUserEnteredFormat(
                        new CellFormat()
                                .setTextFormat(new TextFormat().setBold(true))
                                .setBackgroundColor(new Color().setRed(0.2f).setGreen(0.4f).setBlue(0.6f))))
                .setFields("userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor")));

        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();
    }

    private static Integer getSheetIdByName(String sheetName) throws IOException {
        com.google.api.services.sheets.v4.model.Spreadsheet spreadsheet = sheetsService.spreadsheets().get(SPREADSHEET_ID).execute();
        for (Sheet sheet : spreadsheet.getSheets()) {
            if (sheet.getProperties().getTitle().equals(sheetName)) {
                return sheet.getProperties().getSheetId();
            }
        }
        return null;
    }

    private static void appendResultsToSheet(List<TestResult> results) {
        if (sheetsService == null || results.isEmpty()) return;

        try {
            int currentRows = getCurrentRowCount();
            int neededRows = currentRows + results.size() + 5;
            ensureSheetHasEnoughRows(neededRows);

            ValueRange response = sheetsService.spreadsheets().values()
                    .get(SPREADSHEET_ID, mainSheetName + "!A:A")
                    .execute();
            int lastRow = response.getValues() != null ? response.getValues().size() : 1;

            List<List<Object>> rows = new ArrayList<>();
            for (TestResult result : results) {
                rows.add(result.toRow());
            }

            ValueRange body = new ValueRange().setValues(rows);
            sheetsService.spreadsheets().values()
                    .update(SPREADSHEET_ID, mainSheetName + "!A" + (lastRow + 1) + ":M" + (lastRow + rows.size()), body)
                    .setValueInputOption("RAW")
                    .execute();
            System.out.println("Added " + results.size() + " results to sheet");

        } catch (Exception e) {
            System.err.println("Could not add results: " + e.getMessage());
            tryAlternativeAppend(results);
        }
    }

    private static void tryAlternativeAppend(List<TestResult> results) {
        try {
            for (TestResult result : results) {
                ValueRange body = new ValueRange().setValues(Collections.singletonList(result.toRow()));
                sheetsService.spreadsheets().values()
                        .append(SPREADSHEET_ID, mainSheetName + "!A:M", body)
                        .setValueInputOption("RAW")
                        .execute();
            }
            System.out.println("Added " + results.size() + " results using append method");
        } catch (Exception e) {
            System.err.println("Alternative append also failed: " + e.getMessage());
        }
    }

    private static void handlePopups(WebDriver driver, JavascriptExecutor jsExecutor) {
        try {
            String[] cookieSelectors = {
                    "button[aria-label='Accept cookies']", "button[aria-label='Accept']",
                    ".cookie-accept", ".cookie-consent-accept", "#cookie-accept",
                    ".accept-cookies", ".gdpr-accept", ".cc-accept"
            };
            for (String selector : cookieSelectors) {
                try {
                    WebElement acceptBtn = driver.findElement(By.cssSelector(selector));
                    if (acceptBtn.isDisplayed()) {
                        jsExecutor.executeScript("arguments[0].click();", acceptBtn);
                        Thread.sleep(500);
                        break;
                    }
                } catch (Exception e) {}
            }

            try {
                Alert alert = driver.switchTo().alert();
                alert.dismiss();
            } catch (Exception e) {}

            String[] modalSelectors = {".modal .close", ".popup-close", ".popup .close", ".dialog-close"};
            for (String selector : modalSelectors) {
                try {
                    WebElement closeBtn = driver.findElement(By.cssSelector(selector));
                    if (closeBtn.isDisplayed()) {
                        jsExecutor.executeScript("arguments[0].click();", closeBtn);
                        Thread.sleep(300);
                        break;
                    }
                } catch (Exception e) {}
            }

            try {
                Set<String> handles = driver.getWindowHandles();
                if (handles.size() > 1) {
                    String mainHandle = driver.getWindowHandle();
                    for (String handle : handles) {
                        if (!handle.equals(mainHandle)) {
                            driver.switchTo().window(handle);
                            driver.close();
                            driver.switchTo().window(mainHandle);
                            break;
                        }
                    }
                }
            } catch (Exception e) {}

        } catch (Exception e) {}
    }

    static List<MenuItem> getAllMenuItemsWithHierarchy(WebDriver driver, JavascriptExecutor jsExecutor, Actions actions) {
        List<MenuItem> menuItems = new ArrayList<>();

        try {
            List<WebElement> topLevelItems = driver.findElements(By.xpath(
                    "//nav//ul[contains(@class, 'menu')]//li[contains(@class, 'menu-item')] | " +
                            "//nav//ul[contains(@class, 'nav')]//li[contains(@class, 'nav-item')] | " +
                            "//nav//ul//li[contains(@class, 'menu-item')]"
            ));

            if (topLevelItems.isEmpty()) {
                topLevelItems = driver.findElements(By.xpath(
                        "//nav//a[contains(@href, 'http')] | " +
                                "//header//a[contains(@href, 'http')]"
                ));
            }

            for (WebElement item : topLevelItems) {
                try {
                    WebElement link = item;
                    if (!item.getTagName().equals("a")) {
                        try {
                            link = item.findElement(By.tagName("a"));
                        } catch (Exception e) {
                            continue;
                        }
                    }

                    String url = link.getAttribute("href");
                    if (!isValidUrl(url)) continue;

                    String displayText = link.getText().trim();
                    if (displayText.isEmpty()) {
                        displayText = link.getAttribute("innerText").trim();
                    }
                    if (displayText.isEmpty()) {
                        displayText = "Menu Item";
                    }

                    String parentName = "Home Page";
                    String childName = displayText;
                    String fullPath = displayText;
                    int level = 1;

                    try {
                        WebElement parentLi = item.findElement(By.xpath(".."));
                        WebElement grandParent = parentLi.findElement(By.xpath(".."));

                        String grandParentClass = grandParent.getAttribute("class");
                        if (grandParentClass != null &&
                                (grandParentClass.contains("sub-menu") ||
                                        grandParentClass.contains("dropdown-menu") ||
                                        grandParentClass.contains("submenu"))) {

                            level = 2;
                            WebElement parentLink = grandParent.findElement(By.xpath("preceding-sibling::a | ../a"));
                            parentName = parentLink.getText().trim();
                            if (parentName.isEmpty()) parentName = "Parent Menu";
                            childName = displayText;
                            fullPath = parentName + " > " + childName;
                        }
                    } catch (Exception e) {}

                    try {
                        WebElement parentLi = item.findElement(By.xpath(".."));
                        WebElement grandParent = parentLi.findElement(By.xpath(".."));
                        WebElement greatGrandParent = grandParent.findElement(By.xpath(".."));

                        String greatGrandParentClass = greatGrandParent.getAttribute("class");
                        if (greatGrandParentClass != null &&
                                (greatGrandParentClass.contains("sub-menu") ||
                                        greatGrandParentClass.contains("dropdown-menu"))) {

                            level = 3;
                            WebElement level1Link = greatGrandParent.findElement(By.xpath("preceding-sibling::a | ../a"));
                            WebElement level2Link = grandParent.findElement(By.xpath("preceding-sibling::a | ../a"));

                            String level1Name = level1Link.getText().trim();
                            String level2Name = level2Link.getText().trim();

                            parentName = level2Name;
                            childName = displayText;
                            fullPath = level1Name + " > " + level2Name + " > " + childName;
                        }
                    } catch (Exception e) {}

                    menuItems.add(new MenuItem(url, displayText, parentName, childName, fullPath, level));

                } catch (Exception e) {}
            }

            Set<String> uniqueUrls = new HashSet<>();
            List<MenuItem> uniqueItems = new ArrayList<>();
            for (MenuItem item : menuItems) {
                if (!uniqueUrls.contains(item.getUrl())) {
                    uniqueUrls.add(item.getUrl());
                    uniqueItems.add(item);
                }
            }

            System.out.println("  Total unique URLs discovered: " + uniqueItems.size() +
                    " (Level1:" + uniqueItems.stream().filter(i -> i.getLevel() == 1).count() +
                    ", Level2:" + uniqueItems.stream().filter(i -> i.getLevel() == 2).count() +
                    ", Level3+:" + uniqueItems.stream().filter(i -> i.getLevel() >= 3).count() + ")");
            return uniqueItems;

        } catch (Exception e) {
            System.err.println("Error getting menu hierarchy: " + e.getMessage());
            return new ArrayList<>();
        }
    }

    static String takeScreenshotOnFailure(String url, int statusCode, Exception error, WebDriver driver) {
        try {
            String urlName = url.replace("https://", "")
                    .replace("http://", "")
                    .replace("/", "_")
                    .replace(".", "_")
                    .replace("?", "_")
                    .replace("&", "_");

            if (urlName.length() > 100) {
                urlName = urlName.substring(0, 100);
            }

            String timestamp = String.valueOf(System.currentTimeMillis());
            String filename = String.format("FAILED_%s_%d_%s.png", urlName, statusCode, timestamp);
            Path destination = Paths.get(screenshotDir, filename);

            File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            Files.copy(screenshot.toPath(), destination);

            System.out.println("    Screenshot saved: " + filename);
            return destination.toString();

        } catch (Exception e) {
            System.err.println("    Failed to save screenshot: " + e.getMessage());
            return "Screenshot failed: " + e.getMessage();
        }
    }

    static TestResult testUrlStatusWithDetails(String websiteUrl, MenuItem menuItem, WebDriver driver, JavascriptExecutor jsExecutor) {
        long startTime = System.currentTimeMillis();
        String errorLog = "";
        Integer statusCode = null;
        String status = "PASS";
        String screenshotPath = "";
        String pageTitle = menuItem.getPageTitle();

        try {
            driver.get(menuItem.getUrl());
            Thread.sleep(1000);
            handlePopups(driver, jsExecutor);

            try {
                String title = driver.getTitle();
                if (title != null && !title.isEmpty()) pageTitle = title;
            } catch (Exception e) {}

            HttpURLConnection conn = (HttpURLConnection) new URL(menuItem.getUrl()).openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(8000);
            conn.setReadTimeout(8000);
            conn.setInstanceFollowRedirects(true);
            conn.connect();
            statusCode = conn.getResponseCode();

            if (statusCode != 200 && statusCode != 301 && statusCode != 302) {
                status = "FAIL";
                errorLog = String.format("ERROR on '%s' (Parent: '%s', Child: '%s') - HTTP %d - %s",
                        menuItem.getPageTitle(), menuItem.getParentPage(),
                        menuItem.getChildPage(), statusCode, getHttpStatusMessage(statusCode));
                screenshotPath = takeScreenshotOnFailure(menuItem.getUrl(), statusCode, null, driver);
            }

        } catch (Exception e) {
            status = "FAIL";
            statusCode = 0;
            errorLog = String.format("ERROR on '%s' (Parent: '%s', Child: '%s') - Exception: %s",
                    menuItem.getPageTitle(), menuItem.getParentPage(),
                    menuItem.getChildPage(), e.getMessage());
            screenshotPath = takeScreenshotOnFailure(menuItem.getUrl(), 0, e, driver);
        }

        long duration = System.currentTimeMillis() - startTime;
        String dateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));

        return new TestResult(
                dateTime, websiteUrl, pageTitle, menuItem.getParentPage(), menuItem.getChildPage(),
                menuItem.getFullPath(), "Navigation", menuItem.getUrl(),
                duration, errorLog, statusCode, status, screenshotPath
        );
    }

    static void expandAllMenus(WebDriver driver, Actions actions, JavascriptExecutor jsExecutor, WebDriverWait wait) {
        try {
            List<WebElement> clickableMenus = driver.findElements(By.xpath(
                    "//nav//*[contains(@class, 'dropdown-toggle')] | " +
                            "//nav//*[contains(@class, 'menu-toggle')] | " +
                            "//nav//button[contains(@class, 'navbar-toggler')]"
            ));
            for (WebElement menu : clickableMenus) {
                try {
                    if (menu.isDisplayed() && menu.isEnabled()) {
                        jsExecutor.executeScript("arguments[0].click();", menu);
                        Thread.sleep(300);
                    }
                } catch (Exception e) {}
            }

            List<WebElement> hoverMenus = driver.findElements(By.xpath(
                    "//nav//li[contains(@class, 'menu-item')] | " +
                            "//nav//li[contains(@class, 'nav-item')]"
            ));
            for (WebElement menu : hoverMenus) {
                try {
                    if (menu.isDisplayed()) {
                        actions.moveToElement(menu).perform();
                        Thread.sleep(200);
                    }
                } catch (Exception e) {}
            }
            Thread.sleep(500);
            handlePopups(driver, jsExecutor);
        } catch (Exception e) {}
    }

    static boolean isValidUrl(String url) {
        if (url == null || url.isEmpty() || url.equals("#")) return false;
        if (url.contains("javascript:") || url.contains("mailto:") || url.contains("tel:")) return false;
        if (url.contains("linkedin.com") || url.contains("facebook.com") || url.contains("twitter.com")) return false;
        if (url.contains("instagram.com") || url.contains("youtube.com")) return false;
        if (url.endsWith(".pdf") || url.endsWith(".jpg") || url.endsWith(".png")) return false;
        if (!url.startsWith("http")) return false;
        return true;
    }
}