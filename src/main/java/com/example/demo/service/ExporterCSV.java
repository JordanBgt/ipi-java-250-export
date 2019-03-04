package com.example.demo.service;


import java.io.IOException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public class ExporterCSV<T> {

    private List<String> headers = new ArrayList<>();

    private List<Function<T, String>> functions = new ArrayList<>();

    public void addColumnString(String headerName, Function<T, String> function) {
        headers.add(headerName);
        functions.add(function);
    }

    public void addColumnLong(String headerName, Function<T, Long> function) {
        headers.add(headerName);
        functions.add(function.andThen(integerValue -> integerValue == null ? "" : integerValue.toString()));
    }

    public void addColumnInteger(String headerName, Function<T, Integer> function) {
        headers.add(headerName);
        functions.add(function.andThen(integerValue -> integerValue == null ? "" : integerValue.toString()));
    }

    public void createCSV(Writer printWriter, List<T> objects) throws IOException {
        for (String header : headers) {
            printWriter.write(header);
            printWriter.write(";");
        }
        printWriter.write("\n");
        for (Function<T, String> function : functions) {
            for (T object : objects) {
                String value = function.apply(object);
                printWriter.write(value);
                printWriter.write(";");
            }
            printWriter.write("\n");
        }
    }

}
