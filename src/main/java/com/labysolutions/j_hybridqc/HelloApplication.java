package com.labysolutions.j_hybridqc;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;
import java.util.HashMap;

import com.labysolutions.j_hybridqc.HybridLogic;

public class HelloApplication extends Application {
    @Override
    public void start(Stage stage) throws Exception {

        System.out.println("Hello World");

        var newHybridQc = new ExcelLogicClass();
        newHybridQc.start();
//        hybridQC = new HybridLogic()
//        FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("hello-view.fxml"));
//        Scene scene = new Scene(fxmlLoader.load(), 320, 240);
////        Scene scene = new Scene();
//        stage.setTitle("Hello!");
//        stage.setScene(scene);
//        stage.show();
    }

//    public static void main(String[] args) {
//        launch();
//    }


//    public boolean is_isogram(String string){
//        var dict_letter = new HashMap<Character, Character>();
//
//        for (char c : string.toCharArray()){
//            var exist = dict_letter.get(c);
//
//            if (exist != null) return false;
//            else dict_letter.put(c, c);
//
//        }
//
//        return true;
//    }
}