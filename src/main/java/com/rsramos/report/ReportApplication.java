package com.rsramos.report;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.gson.*;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@SpringBootApplication
public class ReportApplication {

    public static void main(String[] args) throws IOException {
//		SpringApplication.run(ReportApplication.class, args);
        createXLS();
    }

    private static void createXLS() throws IOException {
        JsonObject json = new JsonObject();
        json.add("config", getConfig());
        json.add("data", getData());
        //CreateReport.createXLS(json);
        System.out.println(json);

        getKeysInJsonUsingJsonNodeFieldNames(json.toString()).forEach(System.out::println);
    }

    private static JsonObject getConfig() {
        JsonObject json = new JsonObject();
        json.addProperty("header", "black");
        json.addProperty("body", "white");
        return json;
    }

    private static JsonArray getData() {
        JsonArray jsonArray = new JsonArray();

        JsonObject json1 = new JsonObject();
        json1.addProperty("nome", "rafa");
        json1.addProperty("descricao", "el");
        json1.addProperty("outro", "se");

        JsonObject json2 = new JsonObject();
        json2.addProperty("nome", "bea");
        json2.addProperty("descricao", "triz");
        json2.addProperty("outro", "fra");

        JsonObject json3 = new JsonObject();
        json3.addProperty("nome", "ratriz");
        json3.addProperty("descricao", "sefra");
        json3.addProperty("outro", "rano");

        jsonArray.add(json1);
        jsonArray.add(json2);
        jsonArray.add(json3);

        return jsonArray;
    }

    public static List<String> getKeysInJsonUsingJsonNodeFieldNames(String json) throws IOException, JsonProcessingException {
        ObjectMapper mapper = new ObjectMapper();
        List<String> keys = new ArrayList<>();
        JsonNode jsonNode = mapper.readTree(json);
        Iterator<String> iterator = jsonNode.fieldNames();
        iterator.forEachRemaining(e -> keys.add(e));
        return keys;
    }

}
