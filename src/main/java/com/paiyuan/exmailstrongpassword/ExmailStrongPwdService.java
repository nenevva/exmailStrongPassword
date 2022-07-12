package com.paiyuan.exmailstrongpassword;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.passay.CharacterData;
import org.passay.CharacterRule;
import org.passay.EnglishCharacterData;
import org.passay.PasswordGenerator;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.SpringApplication;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;

import javax.annotation.PostConstruct;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Objects;

@Service
@RequiredArgsConstructor
public class ExmailStrongPwdService {

    private final RestTemplate restTemplate;

    @Value("${corpid}")
    private final String corpid;

    @Value("${corpsecret}")
    private final String corpsecret;

    private String access_token;

    public JSONObject requestSender(String url) {

        String response = restTemplate.getForObject(url, String.class);
        return JSON.parseObject(response);
    }

    public void getAccessToken() {

        String url = "https://api.exmail.qq.com/cgi-bin/gettoken?corpid=" + corpid + "&corpsecret=" + corpsecret;
        JSONObject jsonObject = requestSender(url);
        access_token = jsonObject.getString("access_token");
    }

    public void changePassword(String userid, String password) {

        String url = "https://api.exmail.qq.com/cgi-bin/user/update?access_token" + access_token;
        JSONObject jsonObject = new JSONObject();
        jsonObject.put("userid", userid);
        jsonObject.put("password", password);
        JSONObject resp = JSONObject.parseObject(restTemplate.postForObject(url, jsonObject, String.class));
        System.out.println(userid + resp.getString("errcode") + " " + resp.getString("errmsg"));

        //TODO: 短信通知
    }

    @PostConstruct
    public void startService() throws IOException {

        getAccessToken();

        PasswordGenerator gen = new PasswordGenerator();

        CharacterData lowerCaseChars = EnglishCharacterData.LowerCase;
        CharacterRule lowerCaseRule = new CharacterRule(lowerCaseChars);
        lowerCaseRule.setNumberOfCharacters(2);

        CharacterData upperCaseChars = EnglishCharacterData.UpperCase;
        CharacterRule upperCaseRule = new CharacterRule(upperCaseChars);
        upperCaseRule.setNumberOfCharacters(2);

        CharacterData digitChars = EnglishCharacterData.Digit;
        CharacterRule digitRule = new CharacterRule(digitChars);
        digitRule.setNumberOfCharacters(2);

        Path filePath = Paths.get("C:\\Users\\wpy7634\\Desktop\\exmail.xlsx");
        InputStream is = Files.newInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row.getCell(12).getNumericCellValue() == 1) {

                String userid = row.getCell(0).getStringCellValue();
                String password = gen.generatePassword(8, lowerCaseRule, upperCaseRule, digitRule);
                changePassword(userid, password);
            }
        }
    }
}
