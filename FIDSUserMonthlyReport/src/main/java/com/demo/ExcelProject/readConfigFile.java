package com.demo.ExcelProject;

import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.io.*;
import java.net.URL;
import java.util.Properties;
import java.util.*;

public class readConfigFile {

    public static String main(String args) {

        Properties prop = new Properties();

        //**************************************
        //get value from properties file
        //**************************************

        try{

            String fileName = "app.config";

            InputStream is = new FileInputStream(fileName);

            prop.load(is);

            //System.out.println(args + ": " + prop.getProperty(args));

        } catch (FileNotFoundException ex) {
            ex.printStackTrace();

        } catch (IOException e) {
            e.printStackTrace();
        }

        return prop.getProperty(args);

    }
}
