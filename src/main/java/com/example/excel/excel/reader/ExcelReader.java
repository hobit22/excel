package com.example.excel.excel.reader;


import com.example.excel.excel.error.ExcelReaderErrorField;
import com.example.excel.excel.error.ExcelReaderException;
import com.example.excel.excel.util.ExcelUtils;
import org.apache.poi.ss.usermodel.Row;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * 엑셀 업로드를 유틸리티
 */
public class ExcelReader {
    /**
     * 엑셀 업로드 에러 필드 리스트
     **/
    public static List<ExcelReaderErrorField> errorFieldList;
    /**
     * 엑셀 업로드 HEADER 리스트
     **/
    public static List<String> headerList;
    private static Logger log = LoggerFactory.getLogger(ExcelReader.class);
    /**
     * SAMPLE
     * 사용하기 위해서는 아래와 같이 엑셀 업로드 객체 생성 후 해당 객체
     * from static Constructor 를 생성하고
     * ExcelUtils.setObjectMapping(new Object(), row); 로 리턴 해야 함
     *
     * Sample Product 객체
     */
    /*
    public static Product from(Row row) {
        return ExcelUtils.setObjectMapping(new Product(), row);
    }
    */

    /**
     * 엑셀 파일의 데이터를 읽어서 요청한 오브젝트 타입 리스트에 담아 준다
     *
     * @param multipartFile 업로드한 파일
     * @param clazz         ExcelSheet 를 상속받은 class
     * @param sheetNumber   sheet 가 몇번째 sheet 인지를 나타내는 숫자
     * @return List<T>
     */
    /*
    public static <T extends ExcelAbstractSheet> List<T> getObjectList(final MultipartFile multipartFile,
                                                                       Class<T> clazz, Integer sheetNumber) {

        errorFieldList = new ArrayList<>();
        headerList = new ArrayList<>();

        if (Objects.isNull(multipartFile)) throw new CustomException(ErrorCode.UNKNOWN);

        try {
            log.info("Excel Upload fileInfo :: fileName: {}, fileSize: {} Byte, {} MB", multipartFile.getOriginalFilename(), multipartFile.getSize(), multipartFile.getSize() / 1024 / 1024);
        } catch (Exception e) {
            log.info("Excel Upload fileInfo :: fileName: {}, fileSize: {}", multipartFile.getOriginalFilename(), "비정상 파일 - 파일 사이즈 측정 불가");
        }

        String originalFileName = multipartFile.getOriginalFilename();
        String originalFileExtension = originalFileName.substring(originalFileName.lastIndexOf(".") + 1);
        if (!(originalFileExtension.equals("xlsx") || originalFileExtension.equals("xls")))
            throw new IllegalArgumentException();

        // 엑셀 파일을 Workbook에 담는다
        final Workbook workbook;
        try {
            workbook = WorkbookFactory.create(multipartFile.getInputStream());
        } catch (IOException e) {
            throw new IllegalArgumentException();
        }

        ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
        int headerRowSize = excelSheet.headerRowSize();
        String sheetName = excelSheet.sheetName();

        // 시트 수 (첫번째에만 존재시 0)
        final Sheet sheet = workbook.getSheetAt(sheetNumber);
        if (!sheet.getSheetName().equals(sheetName)) {
            throw new IllegalArgumentException();
        }
        // 전체 행 수
        final int rowCount = sheet.getLastRowNum();

        List<T> objectList = new ArrayList<>();

        try {
            T tmp = clazz.newInstance();

            headerList = getHeader(multipartFile, headerRowSize - 1, sheetNumber);

            for (int i = headerRowSize; i < rowCount + 1; i++) {
                Row targetRow = sheet.getRow(i);

                if (isPass(targetRow)) {
                    objectList.add(tmp.from(targetRow));
                }
            }

        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }

        if (!ListUtils.emptyIfNull(errorFieldList).isEmpty()) {
            throw new ExcelReaderException(errorFieldList);
        }


        return objectList;
    }

     */

    /**
     * 해당 ROW에 있는 index가 비어 있으면 빈 ROW로 판단
     *
     * @param row
     * @return boolean
     */
    private static boolean isPass(Row row) {
        int i = 0;
        boolean isPass = false;
        String cellValue = ExcelUtils.getValue(row.getCell(0));

        isPass = !cellValue.isEmpty() && !cellValue.isBlank();

        // log.debug("## row.getPhysicalNumberOfCells() = {}, isPass = {}",row.getPhysicalNumberOfCells(), isPass);
        return isPass;

    }

    /**
     * 헤더 가져오기
     * 가장 상단에 헤더가 있다면 헤더 정보를 List<String> 에 담아준다
     *
     * @param multipartFile 엑셀파일
     * @param rowNumber     헤더가 있는 row number 값
     * @param sheetNumber   sheet 가 몇번째 sheet 인지를 나타내는 숫자
     * @return List<String> 헤더 리스트
     * @throws CustomException
     */
    /*
    public static List<String> getHeader(final MultipartFile multipartFile, Integer rowNumber, Integer sheetNumber) {

        // rowNumber 이 입력되지 않으면 default로 0 번째 라인을 header 로 판단
        if (Objects.isNull(rowNumber)) rowNumber = 0;

        final Workbook workbook;
        try {
            workbook = WorkbookFactory.create(multipartFile.getInputStream());
        } catch (IOException e) {
//            throw new CustomException(ErrorCode.EXCEL_READER_FAILED);
        }
        // 시트 수 (첫번째에만 존재시 0)
        final Sheet sheet = workbook.getSheetAt(sheetNumber);

        // 타이틀 가져오기
        Row title = sheet.getRow(rowNumber);
        return IntStream
                .range(0, title.getPhysicalNumberOfCells())
                .mapToObj(cellIndex -> title.getCell(cellIndex).getStringCellValue())
                .collect(Collectors.toList());
    }
     */
}
