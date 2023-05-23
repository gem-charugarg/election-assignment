import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.*;

public class Election {
    private static final WebDriver driver = new ChromeDriver();
    private final static Logger logger = LogManager.getLogger(Election.class);

    static int votesLessThanNota = 0, votesGreaterThan50 = 0;

    public static void launchURL(String url) {
        try {
            driver.get(url);
            logger.info("URL launched: " + url);
        } catch (Exception e) {
            logger.error("Unable to launch URL");
        }
    }

    public static void waitForElement(By locator, int duration) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds((long) duration));
        WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(locator));
    }

    public static void selectByIndex(By locator, int index) {
        Select select = new Select(driver.findElement(locator));
        select.selectByIndex(index);
    }

    public static void main(String[] args) throws InterruptedException {
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Charu.garg\\Desktop\\Election\\chromedriver_win32 (3)\\chromedriver.exe");
        driver.manage().window().maximize();
        launchURL("https://results.eci.gov.in/ResultAcGenMar2022/ConstituencywiseS0510.htm?ac=10");
        getCandidateWonMinimumVote();
        getMaxVoteDifference();
        getMinVotesInEachState();
        getCountOfVotesLessThanNota();
        getCountOfCandidatesWithVotesMoreThan50();
        getMaxVotesInEachState();
    }

    public static XSSFWorkbook workbook = new XSSFWorkbook();

    private static void getMaxVoteDifference() {
        waitForElement(locators.stateSelector, 10);
        List<WebElement> stateNames = driver.findElements(locators.allStates);
        List<String> states = new ArrayList<>();
        Map<String, Integer> maxVoteDiffPerState = new HashMap<>();
        Map<String, Double> maxVotePercentageDiffPerState = new HashMap<>();

        XSSFSheet MaxVoteDiffSheet = workbook.createSheet("MaxVoteDiff");
        XSSFSheet MaxVotePercentageDiffPerStateSheet = workbook.createSheet("MaxVotePercentageDiffPerState");
        Map<Integer, Object[]> MaxVoteDiffData = new TreeMap<Integer, Object[]>();
        Map<Integer, Object[]> MaxVotePercentageDiffPerStateData = new TreeMap<Integer, Object[]>();
        //Excel Headers
        int val = 1;
        MaxVoteDiffData.put(val, new Object[]{"State Name", "Max Vote Candidate Name", "Constituencies", "Max Vote Diff"});
        MaxVotePercentageDiffPerStateData.put(val, new Object[]{"State Name", "Max Vote Per Candidate Name", "Constituencies", "Max Vote Percentage Difference Per State"});

        for (int i = 1; i < stateNames.size(); i++) {
            selectByIndex(locators.stateSelector, i);
            states.add(getSelectedOption(locators.stateSelector));
            List<WebElement> constituencyNames = driver.findElements(locators.allConstituencies);
            List<String> constituencies = new ArrayList<>();
            List<String> maxVoteCandidate = new ArrayList<>();
            List<Integer> maxVoteDiff = new ArrayList<>();
            List<Double> maxVotePercentageDiff = new ArrayList<>();
            List<String> maxVotePercentageCandidate = new ArrayList<>();
            List<Integer> maxVotes = new ArrayList<>();
            List<Double> maxVotePercentage = new ArrayList<>();
            for (int j = 1; j < constituencyNames.size(); j++) {
                selectByIndex(locators.constituencySelector, j);
                constituencies.add(getSelectedOption(locators.constituencySelector));
                //getColIndex for taking particular column in table.
                int totalVotesColumnIndex = getColIndex("Total Votes");
                int percentageColumnIndex = getColIndex("% of Votes");
                int candidateColIndex = getColIndex("Candidate");
                waitForElement(locators.tableRows, 10);
                int totalRows = driver.findElements(locators.tableRows).size();
                int maximum = -10;
                int secondMax = Integer.MIN_VALUE;
                Double max = -10.00;
                Double secondMaximum = Double.MIN_VALUE;
                String candidateName = new String();
                String candidateNameForPercentage = new String();
                for (int k = 4; k < totalRows - 1; k++) {
                    int current = Integer.parseInt(driver.findElement(locators.colName(k, totalVotesColumnIndex)).getText());
                    double curr = Double.parseDouble(driver.findElement(locators.colName(k, percentageColumnIndex)).getText());
                    if (maximum < current) {
                        secondMax = maximum;
                        maximum = current;
                        //candidateName = new String();
                        candidateName = driver.findElement(locators.colName(k, candidateColIndex)).getText();
                    } else if (current > secondMax && current < maximum) {
                        secondMax = current;
                    }

                    if (max < curr) {
                        secondMaximum = max;
                        max = curr;
                        //candidateNameForPercentage = new String();
                        candidateNameForPercentage = driver.findElement(locators.colName(k, candidateColIndex)).getText();
                    } else if (curr > secondMaximum && curr < max) {
                        secondMaximum = curr;
                    }
                }
                maxVoteCandidate.add(candidateName);
                maxVotes.add(maximum);
                maxVoteDiff.add(maximum - secondMax);
                maxVotePercentageDiff.add(max - secondMaximum);
                maxVotePercentage.add(max);
                maxVotePercentageCandidate.add(candidateNameForPercentage);
            }
            int maxVoteDiffIndex = getMaxVoteIndex(maxVoteDiff);
            int maxVotePercentageDiffIndex = getMaxVotePercentageIndex(maxVotePercentageDiff);
            maxVoteDiffPerState.put(maxVoteCandidate.get(maxVoteDiffIndex), maxVoteDiff.get(maxVoteDiffIndex));
            maxVotePercentageDiffPerState.put(maxVotePercentageCandidate.get(maxVotePercentageDiffIndex), maxVotePercentageDiff.get(maxVotePercentageDiffIndex));
            logger.info(states.get(i - 1) + "::::" + maxVoteCandidate.get(maxVoteDiffIndex) + "::::" + maxVoteDiff.get(maxVoteDiffIndex) + "::::" + constituencies.get(maxVoteDiffIndex));
            logger.info(states.get(i - 1) + "::::" + maxVotePercentageCandidate.get(maxVotePercentageDiffIndex) + "::::" + maxVotePercentageDiff.get(maxVotePercentageDiffIndex) + "::::" + constituencies.get(maxVotePercentageDiffIndex));
            val++;
            //print value in excel
            MaxVoteDiffData.put(val, new Object[]{states.get(i - 1), maxVoteCandidate.get(maxVoteDiffIndex), constituencies.get(maxVoteDiffIndex), maxVoteDiff.get(maxVoteDiffIndex)});
            MaxVotePercentageDiffPerStateData.put(val, new Object[]{states.get(i - 1), maxVotePercentageCandidate.get(maxVotePercentageDiffIndex) + "", constituencies.get(maxVotePercentageDiffIndex), maxVotePercentageDiff.get(maxVotePercentageDiffIndex)});

        }
        //toPrint Data in Excel
        iterateExcel(MaxVoteDiffData, workbook, MaxVoteDiffSheet);
        iterateExcel(MaxVotePercentageDiffPerStateData, workbook, MaxVotePercentageDiffPerStateSheet);

    }

    private static void getMaxVotesInEachState() throws InterruptedException {
        waitForElement(locators.stateSelector, 10);
        List<WebElement> stateNames = driver.findElements(locators.allStates);
        List<String> states = new ArrayList<>();
        Map<String, Integer> maxVotesPerState = new HashMap<>();
        Map<String, Double> maxVotePercentagePerState = new HashMap<>();
        XSSFSheet maxVotesPerStateSheet = workbook.createSheet("Max Vote In Each State");
        XSSFSheet maxVotePercentagePerStateSheet = workbook.createSheet("Max Vote Percentage In Each State");
        Map<Integer, Object[]> maxVotesPerStateData = new TreeMap<Integer, Object[]>();
        Map<Integer, Object[]> maxVotePercentagePerStateData = new TreeMap<Integer, Object[]>();
        //Excel Headers
        int val = 1;
        maxVotesPerStateData.put(val, new Object[]{"State Name", "MaxVoteCandidateName", "Constituencies", "Max Vote In Each State"});
        maxVotePercentagePerStateData.put(val, new Object[]{"State Name", "MaxVotePerCandidateName", "Constituencies", "Max Vote Percentage In Each State"});

        for (int i = 1; i < stateNames.size(); i++) {
            selectByIndex(locators.stateSelector, i);
            states.add(getSelectedOption(locators.stateSelector));
            List<WebElement> constituencyNames = driver.findElements(locators.allConstituencies);
            List<String> constituencies = new ArrayList<>();
            List<String> maxVoteCandidate = new ArrayList<>();
            List<String> maxVotePercentageCandidate = new ArrayList<>();
            List<Integer> maxVotes = new ArrayList<>();
            List<Double> maxVotePercentage = new ArrayList<>();
            for (int j = 1; j < constituencyNames.size(); j++) {
                selectByIndex(locators.constituencySelector, j);
                constituencies.add(getSelectedOption(locators.constituencySelector));
                int totalVotesColumnIndex = getColIndex("Total Votes");
                int percentageColumnIndex = getColIndex("% of Votes");
                int candidateColIndex = getColIndex("Candidate");
                waitForElement(locators.tableRows, 10);
                int totalRows = driver.findElements(locators.tableRows).size();
                int maximum = Integer.MIN_VALUE;
                Double max = Double.MIN_VALUE;
                String candidateName = new String();
                String candidateNameForPercentage = new String();
                for (int k = 4; k < totalRows - 1; k++) {
                    int current = Integer.parseInt(driver.findElement(locators.colName(k, totalVotesColumnIndex)).getText());
                    double curr = Double.parseDouble(driver.findElement(locators.colName(k, percentageColumnIndex)).getText());
                    if (maximum < current) {
                        maximum = current;
                        //candidateName = new String();
                        candidateName = driver.findElement(locators.colName(k, candidateColIndex)).getText();
                    }
                    if (max < curr) {
                        max = curr;
                        //candidateNameForPercentage = new String();
                        candidateNameForPercentage = driver.findElement(locators.colName(k, candidateColIndex)).getText();
                    }
                }
                maxVoteCandidate.add(candidateName);
                maxVotes.add(maximum);
                maxVotePercentage.add(max);
                maxVotePercentageCandidate.add(candidateNameForPercentage);
            }
            int maxVoteIndex = getMaxVoteIndex(maxVotes);
            int maxVotePercentageIndex = getMaxVotePercentageIndex(maxVotePercentage);
            maxVotesPerState.put(maxVoteCandidate.get(maxVoteIndex), maxVotes.get(maxVoteIndex));
            maxVotePercentagePerState.put(maxVotePercentageCandidate.get(maxVotePercentageIndex), maxVotePercentage.get(maxVotePercentageIndex));
            logger.info(states.get(i - 1) + ":::::" + maxVoteCandidate.get(maxVoteIndex) + ":::::" + maxVotes.get(maxVoteIndex) + ":::::" + constituencies.get(maxVoteIndex));
            logger.info(states.get(i - 1) + ":::::" + maxVotePercentageCandidate.get(maxVoteIndex) + ":::::" + maxVotePercentage.get(maxVoteIndex) + ":::::" + constituencies.get(maxVotePercentageIndex));
            val++;
            //print value in excel
            maxVotesPerStateData.put(val, new Object[]{states.get(i - 1), maxVoteCandidate.get(maxVoteIndex), constituencies.get(maxVoteIndex), maxVotes.get(maxVoteIndex)});
            maxVotePercentagePerStateData.put(val, new Object[]{states.get(i - 1), maxVotePercentageCandidate.get(maxVoteIndex) + "", constituencies.get(maxVotePercentageIndex), maxVotePercentage.get(maxVoteIndex)});

        }
        //toPrint Data in Excel
        iterateExcel(maxVotesPerStateData, workbook, maxVotesPerStateSheet);
        iterateExcel(maxVotePercentagePerStateData, workbook, maxVotePercentagePerStateSheet);
    }

    private static void getCountOfCandidatesWithVotesMoreThan50() throws InterruptedException {
        waitForElement(locators.stateSelector, 10);
        List<String> states = new ArrayList<>();
        List<WebElement> stateNames = driver.findElements(locators.allStates);
        Map<String, Integer> countOfCandidatesWithVotesMoreThan50 = new HashMap<>();
        XSSFSheet countOfCandidatesWithVotesMoreThan50DataSheet = workbook.createSheet("Candidate Greater Than 50% Votes");
        Map<Integer, Object[]> countOfCandidatesWithVotesMoreThan50Data = new TreeMap<Integer, Object[]>();
        //Excel Headers
        int val = 1;
        countOfCandidatesWithVotesMoreThan50Data.put(val, new Object[]{"State Name", "Candidate Greater Than 50% Votes"});

        for (int i = 1; i < stateNames.size(); i++) {
            selectByIndex(locators.stateSelector, i);
            states.add(getSelectedOption(locators.stateSelector));
            List<WebElement> constituencyNames = driver.findElements(locators.allConstituencies);
            for (int j = 1; j < constituencyNames.size(); j++) {
                selectByIndex(locators.constituencySelector, j);
                int percentageColumnIndex = Election.getColIndex("% of Votes");
                waitForElement(locators.tableRows, 10);
                int totalRows = driver.findElements(locators.tableRows).size();
                for (int k = 4; k < totalRows - 1; k++) {
                    double current = Double.parseDouble(driver.findElement(locators.colName(k, percentageColumnIndex)).getText());
                    if (current > 50) {
                        votesGreaterThan50++;
                    }
                }
            }
            countOfCandidatesWithVotesMoreThan50.put(states.get(i - 1), votesGreaterThan50);
            logger.info(votesGreaterThan50);
            val++;
            countOfCandidatesWithVotesMoreThan50Data.put(val, new Object[]{states.get(i - 1), votesGreaterThan50});

        }
        iterateExcel(countOfCandidatesWithVotesMoreThan50Data, workbook, countOfCandidatesWithVotesMoreThan50DataSheet);
    }

    private static void getCountOfVotesLessThanNota() throws InterruptedException {
        waitForElement(locators.stateSelector, 10);
        List<String> states = new ArrayList<>();
        List<WebElement> stateNames = driver.findElements(locators.allStates);
        Map<String, Integer> countOfVotesLessThanNota = new HashMap<>();
        XSSFSheet countOfVotesLessThanNotaSheet = workbook.createSheet("GetCountOfVotesLessThanNota");
        Map<Integer, Object[]> countOfVotesLessThanNotaData = new TreeMap<Integer, Object[]>();
        //Excel Headers
        int val = 1;
        countOfVotesLessThanNotaData.put(val, new Object[]{"State Name", "Count Of Votes Less Than Nota"});
        for (int i = 1; i < stateNames.size(); i++) {
            selectByIndex(locators.stateSelector, i);
            states.add(getSelectedOption(locators.stateSelector));
            List<WebElement> constituencyNames = driver.findElements(locators.allConstituencies);
            for (int j = 1; j < constituencyNames.size(); j++) {
                selectByIndex(locators.constituencySelector, j);
                int totalVotesColumnIndex = Election.getColIndex("Total Votes");
                int candidateCol = Election.getColIndex("Candidate");
                waitForElement(locators.tableRows, 10);
                int totalRows = driver.findElements(locators.tableRows).size();
                int notaVotesIndex = getRowIndexOfNota(totalRows, candidateCol);
                int notaVotes = Integer.parseInt(driver.findElement(locators.colName(notaVotesIndex, totalVotesColumnIndex)).getText());
                for (int k = 4; k < totalRows - 1; k++) {
                    int current = Integer.parseInt(driver.findElement(locators.colName(k, totalVotesColumnIndex)).getText());
                    if (k != notaVotesIndex) {
                        if (current < notaVotes) {
                            votesLessThanNota++;
                        }
                    }
                }
            }
            countOfVotesLessThanNota.put(states.get(i - 1), votesLessThanNota);
            logger.info(votesLessThanNota);
            val++;
            countOfVotesLessThanNotaData.put(val, new Object[]{states.get(i - 1), votesLessThanNota});
        }
        iterateExcel(countOfVotesLessThanNotaData, workbook, countOfVotesLessThanNotaSheet);
    }

    private static void getCandidateWonMinimumVote() throws InterruptedException {
        waitForElement(locators.stateSelector, 10);
        List<WebElement> stateNames = driver.findElements(locators.allStates);
        Map<String, Integer> minVoteDiffPerState = new HashMap<>();
        Map<String, Double> minVotePercentageDiffPerState = new HashMap<>();
        List<String> states = new ArrayList<>();
//        Map<String, Integer> maxVotesPerState = new HashMap<>();
//        Map<String, Integer> minVotesCollection = new HashMap<>();
//        Map<String, Double> maxVotePercentagePerState = new HashMap<>();
        XSSFSheet minVoteDiffPerStateSheet = workbook.createSheet("Min Vote Diff In Each State");
        XSSFSheet minVotePercentageDiffPerStateSheet = workbook.createSheet("Min Vote Percentage Diff In Each State");
        Map<Integer, Object[]> minVoteDiffPerStateData = new TreeMap<Integer, Object[]>();
        Map<Integer, Object[]> minVotePercentageDiffPerStateData = new TreeMap<Integer, Object[]>();
        //Excel Headers
        int val = 1;
        minVoteDiffPerStateData.put(val, new Object[]{"State Name", "Max Vote Candidate Name", "Constituencies", "Min Vote Diff In Each State"});
        minVotePercentageDiffPerStateData.put(val, new Object[]{"State Name", "Max Vote Per Candidate Name", "Constituencies", "Min Vote Percentage Diff In Each State"});

        for (int i = 1; i < stateNames.size(); i++) {
            selectByIndex(locators.stateSelector, i);
            states.add(getSelectedOption(locators.stateSelector));
            List<WebElement> constituencyNames = driver.findElements(locators.allConstituencies);
            List<String> constituencies = new ArrayList<>();
            List<String> maxVoteCandidate = new ArrayList<>();
            List<Integer> maxVoteDiff = new ArrayList<>();
            List<Double> maxVotePercentageDiff = new ArrayList<>();
            List<String> maxVotePercentageCandidate = new ArrayList<>();
            List<Integer> maxVotes = new ArrayList<>();
            List<Double> maxVotePercentage = new ArrayList<>();
            for (int j = 1; j < constituencyNames.size(); j++) {
                selectByIndex(locators.constituencySelector, j);
                constituencies.add(getSelectedOption(locators.constituencySelector));
                //getColIndex for taking particular column in table.
                int totalVotesColumnIndex = getColIndex("Total Votes");
                int percentageColumnIndex = getColIndex("% of Votes");
                int candidateColIndex = getColIndex("Candidate");
                waitForElement(locators.tableRows, 10);
                int totalRows = driver.findElements(locators.tableRows).size();
                int maximum = -10;
                int secondMax = Integer.MIN_VALUE;
                Double max = -10.00;
                Double secondMaximum = Double.MIN_VALUE;
                String candidateName = new String();
                String candidateNameForPercentage = new String();
                for (int k = 4; k < totalRows - 1; k++) {
                    int current = Integer.parseInt(driver.findElement(locators.colName(k, totalVotesColumnIndex)).getText());
                    double curr = Double.parseDouble(driver.findElement(locators.colName(k, percentageColumnIndex)).getText());
                    if (maximum < current) {
                        secondMax = maximum;
                        maximum = current;
                        //candidateName = new String();
                        candidateName = driver.findElement(locators.colName(k, candidateColIndex)).getText();
                    } else if (current > secondMax && current < maximum) {
                        secondMax = current;
                    }

                    if (max < curr) {
                        secondMaximum = max;
                        max = curr;
                        //candidateNameForPercentage = new String();
                        candidateNameForPercentage = driver.findElement(locators.colName(k, candidateColIndex)).getText();
                    } else if (curr > secondMaximum && curr < max) {
                        secondMaximum = curr;
                    }
                }
                maxVoteCandidate.add(candidateName);
                maxVotes.add(maximum);
                maxVoteDiff.add(maximum - secondMax);
                maxVotePercentageDiff.add(max - secondMaximum);
                maxVotePercentage.add(max);
                maxVotePercentageCandidate.add(candidateNameForPercentage);
            }
            int minVoteDiffIndex = getMinVoteIndex(maxVoteDiff);
            int minVotePercentageDiffIndex = getMinVotePercentageIndex(maxVotePercentageDiff);
            minVoteDiffPerState.put(maxVoteCandidate.get(minVoteDiffIndex), maxVoteDiff.get(minVoteDiffIndex));
            minVotePercentageDiffPerState.put(maxVotePercentageCandidate.get(minVotePercentageDiffIndex), maxVotePercentageDiff.get(minVotePercentageDiffIndex));
            logger.info(states.get(i - 1) + "::::" + maxVoteCandidate.get(minVoteDiffIndex) + "::::" + maxVoteDiff.get(minVoteDiffIndex) + "::::" + constituencies.get(minVoteDiffIndex));
            logger.info(states.get(i - 1) + "::::" + maxVotePercentageCandidate.get(minVotePercentageDiffIndex) + "::::" + maxVotePercentageDiff.get(minVotePercentageDiffIndex) + "::::" + constituencies.get(minVotePercentageDiffIndex));
            val++;
            //print value in excel
            minVoteDiffPerStateData.put(val, new Object[]{states.get(i - 1), maxVoteCandidate.get(minVoteDiffIndex), constituencies.get(minVoteDiffIndex), maxVoteDiff.get(minVoteDiffIndex)});
            minVotePercentageDiffPerStateData.put(val, new Object[]{states.get(i - 1), maxVotePercentageCandidate.get(minVotePercentageDiffIndex) + "", constituencies.get(minVotePercentageDiffIndex), maxVotePercentageDiff.get(minVotePercentageDiffIndex)});

        }
        //toPrint Data in Excel
        iterateExcel(minVoteDiffPerStateData, workbook, minVoteDiffPerStateSheet);
        iterateExcel(minVotePercentageDiffPerStateData, workbook, minVotePercentageDiffPerStateSheet);
    }

    private static void getMinVotesInEachState() throws InterruptedException {
        waitForElement(locators.stateSelector, 10);
        List<WebElement> stateNames = driver.findElements(locators.allStates);
        List<String> states = new ArrayList<>();
        Map<String, Integer> maxVotesPerState = new HashMap<>();
        XSSFSheet maxVotesPerStateSheet = workbook.createSheet("Min Vote In Each State");
        Map<Integer, Object[]> maxVotesPerStateData = new TreeMap<Integer, Object[]>();
        //Excel Headers
        int val = 1;
        maxVotesPerStateData.put(val, new Object[]{"State Name", "Min Vote Candidate Name", "Constituencies", "Min Vote In Each State"});
        for (int i = 1; i < stateNames.size(); i++) {
            selectByIndex(locators.stateSelector, i);
            states.add(getSelectedOption(locators.stateSelector));
            List<WebElement> constituencyNames = driver.findElements(locators.allConstituencies);
            List<String> constituencies = new ArrayList<>();
            List<String> maxVoteCandidate = new ArrayList<>();
            List<String> maxVotePercentageCandidate = new ArrayList<>();
            List<Integer> maxVotes = new ArrayList<>();
            List<Double> maxVotePercentage = new ArrayList<>();
            for (int j = 1; j < constituencyNames.size(); j++) {
                selectByIndex(locators.constituencySelector, j);
                constituencies.add(getSelectedOption(locators.constituencySelector));
                int totalVotesColumnIndex = getColIndex("Total Votes");
                int percentageColumnIndex = getColIndex("% of Votes");
                int candidateColIndex = getColIndex("Candidate");
                waitForElement(locators.tableRows, 10);
                int totalRows = driver.findElements(locators.tableRows).size();
                int maximum = Integer.MIN_VALUE;
                Double max = Double.MIN_VALUE;
                String candidateName = new String();
                String candidateNameForPercentage = new String();
                for (int k = 4; k < totalRows - 1; k++) {
                    int current = Integer.parseInt(driver.findElement(locators.colName(k, totalVotesColumnIndex)).getText());
                    double curr = Double.parseDouble(driver.findElement(locators.colName(k, percentageColumnIndex)).getText());
                    if (maximum < current) {
                        maximum = current;
                        //candidateName = new String();
                        candidateName = driver.findElement(locators.colName(k, candidateColIndex)).getText();
                    }
                    if (max < curr) {
                        max = curr;
                        //candidateNameForPercentage = new String();
                        candidateNameForPercentage = driver.findElement(locators.colName(k, candidateColIndex)).getText();
                    }
                }
                maxVoteCandidate.add(candidateName);
                maxVotes.add(maximum);
                maxVotePercentage.add(max);
                maxVotePercentageCandidate.add(candidateNameForPercentage);
            }
            int minVoteIndex = getMinVoteIndex(maxVotes);
            maxVotesPerState.put(maxVoteCandidate.get(minVoteIndex), maxVotes.get(minVoteIndex));
            logger.info(states.get(i - 1) + ":::::" + maxVoteCandidate.get(minVoteIndex) + ":::::" + maxVotes.get(minVoteIndex) + ":::::" + constituencies.get(minVoteIndex));
            val++;
            //print value in excel
            maxVotesPerStateData.put(val, new Object[]{states.get(i - 1), maxVoteCandidate.get(minVoteIndex), constituencies.get(minVoteIndex), maxVotes.get(minVoteIndex)});
        }
        //toPrint Data in Excel
        iterateExcel(maxVotesPerStateData, workbook, maxVotesPerStateSheet);
    }

    private static int getMinVoteIndex(List<Integer> minVotes) {
        int min = Integer.MAX_VALUE, index = 0;
        for (int i = 0; i < minVotes.size(); i++) {
            if (min > minVotes.get(i)) {
                min = minVotes.get(i);
                index = i;
            }
        }
        return index;
    }

    private static int getRowIndexOfNota(int totalRows, int candidateCol) {
        int k = 0;
        for (int i = 4; i < totalRows - 1; i++) {
            if (driver.findElement(locators.colName(i, candidateCol)).getText().contains("NOTA")) {
                k = i;
                break;
            }
        }
        return k;
    }

    private static int getMaxVoteIndex(List<Integer> maxVotes) {
        int max = Integer.MIN_VALUE, index = 0;
        for (int i = 0; i < maxVotes.size(); i++) {
            if (max < maxVotes.get(i)) {
                max = maxVotes.get(i);
                index = i;
            }
        }
        return index;
    }

    private static int getMaxVotePercentageIndex(List<Double> maxVotes) {
        double max = Double.MIN_VALUE;
        int index = 0;
        for (int i = 0; i < maxVotes.size(); i++) {
            if (max < maxVotes.get(i)) {
                max = maxVotes.get(i);
                index = i;
            }
        }
        return index;
    }

    private static int getMinVotePercentageIndex(List<Double> maxVotes) {
        double max = Double.MAX_VALUE;
        int index = 0;
        for (int i = 0; i < maxVotes.size(); i++) {
            if (max > maxVotes.get(i)) {
                max = maxVotes.get(i);
                index = i;
            }
        }
        return index;
    }

    private static String getSelectedOption(By locator) {
        Select select = new Select(driver.findElement(locator));
        WebElement element = select.getFirstSelectedOption();
        return element.getText();
    }

    static int getColIndex(String colName) {
        int col = 0;
        List<WebElement> cols = driver.findElements(locators.tableCols);
        for (int i = 0; i < cols.size(); i++) {
            if (cols.get(i).getText().contains(colName)) {
                col = i;
                break;
            }
        }
        return col + 1;
    }

    private static void iterateExcel(Map<Integer, Object[]> data, XSSFWorkbook workbook, XSSFSheet sheet) {
        Set<Integer> keyset = data.keySet();
        int rownum = 0;
        for (Integer key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
                else if (obj instanceof Double)
                    cell.setCellValue((Double) obj);
            }
        }
        try {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("ans1.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

}
