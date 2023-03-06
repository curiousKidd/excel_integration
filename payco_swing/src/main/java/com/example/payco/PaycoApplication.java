package com.example.payco;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.swing.*;

@SpringBootApplication
public class PaycoApplication extends JFrame {

    public static void main(String[] args) {
        SpringApplication.run(PaycoApplication.class, args);
        new FrameSetting();
    }

}


