package br.com.process.excel.endpoint.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class BookResponse implements Serializable {
    private String title;
    private String author;
    private String pages;
}
