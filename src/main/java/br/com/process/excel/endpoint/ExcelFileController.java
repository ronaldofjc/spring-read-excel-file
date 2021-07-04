package br.com.process.excel.endpoint;

import br.com.process.excel.endpoint.entity.BookResponse;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@Slf4j
@RestController
@Tag(name = "Process Excel File")
public class ExcelFileController {

    @ResponseStatus(HttpStatus.CREATED)
    @Operation(summary = "Endpoint to receive excel file")
    @PostMapping(value = "/files/excel", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE, MediaType.MULTIPART_FORM_DATA_VALUE})
    public List<BookResponse> uploadFile(@RequestParam("file") final MultipartFile file) throws IOException {
        final String fileName = StringUtils.cleanPath(Objects.requireNonNull(file.getOriginalFilename()));
        log.info("Filename: {}. Type: {}", fileName, file.getContentType());

        List<BookResponse> books = new ArrayList<>();

        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        XSSFSheet workSheet = workbook.getSheetAt(0);

        for (int i=1;i<workSheet.getPhysicalNumberOfRows();i++) {
            XSSFRow row = workSheet.getRow(i);
            books.add(BookResponse.builder()
                    .title(String.valueOf(row.getCell(0)))
                    .author(String.valueOf(row.getCell(1)))
                    .pages(String.valueOf(row.getCell(2)))
                    .build());
        }

        log.info("Rows: {}", books.size());
        return books;
    }

}
