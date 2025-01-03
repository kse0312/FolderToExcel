package com.ns.foldertoexcel;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.DirectoryChooser;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class MainController {

    // FXML 요소들에 대한 참조
    @FXML
    private TextField folderPathField;

    @FXML
    public Button folderSelectionButton;

    @FXML
    private ListView<String> fileListView;

    @FXML
    private TextField excelFileNameField;

    @FXML
    private ComboBox<Integer> startRowComboBox;

    @FXML
    private ComboBox<String> startColumnComboBox;

    @FXML
    private TextField downloadPathField;

    // 폴더 경로 선택 버튼 클릭 시 실행되는 메소드
    @FXML
    private void handleFolderSelection() {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("폴더 선택");

        // 현재 작업 디렉토리를 초기 값으로 설정
        File defaultDirectory = new File(System.getProperty("user.home"));
        directoryChooser.setInitialDirectory(defaultDirectory);

        // 폴더 선택 창을 띄우고, 사용자가 폴더를 선택하면 해당 경로를 TextField 에 표시
        File selectedDirectory = directoryChooser.showDialog(folderSelectionButton.getScene().getWindow());

        if (selectedDirectory != null) {
            folderPathField.setText(selectedDirectory.getAbsolutePath());
            loadFolderList(selectedDirectory); // 폴더를 선택했으면 해당 폴더 내 파일 리스트를 로드
        }
    }

    // 폴더 내 파일 목록을 ListView 에 로드하는 메소드
    private void loadFolderList(File folder) {
        fileListView.getItems().clear(); // 기존 파일 목록을 지운다.

        // 폴더 내 파일들만 가져오기
        try {
            Files.list(folder.toPath())
                    .filter(Files::isDirectory)  // 디렉토리만 필터링
                    .map(path -> path.getFileName().toString())  // 폴더명만 가져옴
                    .forEach(folderName -> {
                        Platform.runLater(() -> fileListView.getItems().add(folderName));  // UI 스레드에서 ListView 에 추가
                    });
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 엑셀로 내보내기 버튼 클릭 시 실행되는 메소드
    @FXML
    private void handleExportToExcel() {
        String folderPath = folderPathField.getText();
        String excelFileName = excelFileNameField.getText();
        String downloadPath = downloadPathField.getText();
        String startColumn = startColumnComboBox.getValue() == null ? "A" : startColumnComboBox.getValue();
        int startRow = startRowComboBox.getValue() == null ? 0 : startRowComboBox.getValue();

        // ListView 의 항목이 비어 있으면 경고 알림을 띄운다
        if (fileListView.getItems().isEmpty()) {
            // 경고 알림 생성
            Alert alert = new Alert(Alert.AlertType.WARNING);
            alert.setTitle("경고");
            alert.setHeaderText(null);
            alert.setContentText("저장할 파일이 없습니다. 폴더 내 파일을 선택해주세요.");
            alert.showAndWait();
            return;  // 더 이상 진행하지 않음
        }

        // 엑셀 파일명이 비어 있으면 기본값 설정
        if (excelFileName == null || excelFileName.trim().isEmpty()) {
            String folderName = "폴더명";

            // 현재 날짜를 yyyyMMdd 형식으로 얻기
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
            String currentDate = sdf.format(new Date());

            // 기본 파일명 생성
            excelFileName = folderName + "_" + currentDate;
        }
        // 파일명에 사용할 수 없는 특수문자 제거
        excelFileName = excelFileName.replaceAll("[/:*?\"<>|]", "_");  // 특수문자는 "_"로 대체
        // 확장자 추가
        excelFileName += ".xlsx";

        // 엑셀 내보내기 로직 구현
        // 예시로 텍스트 출력
        System.out.println("엑셀로 내보내기...");
        System.out.println("폴더 경로: " + folderPath);
        System.out.println("엑셀 파일명: " + excelFileName);
        System.out.println("시작 행: " + startRow);
        System.out.println("시작 열: " + startColumn);
        System.out.println("다운로드 경로: " + downloadPath);

        // 엑셀 파일 생성
        Workbook workbook = new XSSFWorkbook(); // 새로운 엑셀 파일 생성
        Sheet sheet = workbook.createSheet("파일 목록");

        // 시작 행 및 열 설정
        int rowNum = (startRow != 0) ? startRow - 1 : 0; // 시작 행 (0부터 시작)
        int colNum = (startColumn != null) ? (startColumn.charAt(0) - 'A') : 0; // 시작 열 (A부터 시작)

        // 파일 목록을 첫 번째 행에 추가
        Row headerRow = sheet.createRow(rowNum);
        Cell headerCell = headerRow.createCell(colNum);
        headerCell.setCellValue("폴더명");

        // ListView 에서 파일 목록 가져오기
        List<String> files = fileListView.getItems();

        // 파일 목록을 엑셀 시트에 추가
        for (int i = 0; i < files.size(); i++) {
            Row row = sheet.createRow(i + rowNum + 1); // 데이터를 추가할 행
            Cell cell = row.createCell(colNum); // 열에 데이터를 삽입
            cell.setCellValue(files.get(i));
        }

        // 열 너비 자동 맞추기
        int maxLength = 0;
        int ColumnNum = ((startColumn != null ? startColumn.charAt(0) : 'A') - 'A');
        for (int rowIndex = startRow - 1; rowIndex < files.size(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(ColumnNum);

            if (cell != null) {
                String cellValue = cell.getStringCellValue();
                maxLength = Math.max(maxLength, cellValue.length());
            }
        }
        sheet.setColumnWidth(ColumnNum, maxLength * 256);

        // 엑셀 파일 저장
        try {
            // 중복 파일명을 피하기 위한 코드
            File file = new File(downloadPath, excelFileName);
            String baseName = excelFileName.substring(0, excelFileName.lastIndexOf('.'));
            String extension = excelFileName.substring(excelFileName.lastIndexOf('.'));

            int counter = 0;

            // 파일이 이미 존재하는지 확인하고, 중복된 파일명이 있다면 "(0)", "(1)" 추가
            while (file.exists()) {
                counter++;
                excelFileName = baseName + "(" + counter + ")" + extension;
                file = new File(downloadPath, excelFileName); // 파일 경로 업데이트
            }

            // 파일로 저장
            try (FileOutputStream fileOut = new FileOutputStream(file)) {
                workbook.write(fileOut);
            }

            workbook.close();
            System.out.println("엑셀 파일이 저장되었습니다: " + file.getAbsolutePath());

            // 엑셀 내보내기 성공 메시지
            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            alert.setTitle("성공");
            alert.setHeaderText(null);
            alert.setContentText("엑셀 파일이 성공적으로 저장되었습니다.");
            alert.showAndWait();
        } catch (IOException e) {
            e.printStackTrace();
            // 오류 발생 시 메시지 출력
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("오류");
            alert.setHeaderText(null);
            alert.setContentText("엑셀 파일을 저장하는 중 오류가 발생했습니다.");
            alert.showAndWait();
        }
    }

    // 초기화 시 콤보박스와 기타 설정을 초기화하는 메소드
    @FXML
    public void initialize() {
        // 시작 행 콤보박스에 1부터 10까지 숫자 추가
        for (int i = 1; i <= 10; i++) {
            startRowComboBox.getItems().add(i);
        }
        startRowComboBox.setValue(1); // 기본값 설정

        // 시작 열 콤보박스에 A~J 열 문자 추가
        startColumnComboBox.getItems().addAll("A", "B", "C", "D", "E", "F", "G", "H", "I", "J");
        startColumnComboBox.setValue("A"); // 기본값 설정

        // 다운로드 경로는 기본적으로 사용자의 홈 디렉토리로 설정
        String userHome = System.getProperty("user.home");  // 사용자 홈 디렉토리
        String downloadsPath = userHome + File.separator + "Downloads";  // 다운로드 폴더 경로

        downloadPathField.setText(downloadsPath);  // 다운로드 폴더 경로를 텍스트 필드에 설정
    }
}