package com.visoft.example;

import java.util.List;

public class TestObject {
    private String conclusionsOfTest;
    private String correctiveAction;

    public String getConclusionsOfTest() {
        return conclusionsOfTest;
    }

    public void setConclusionsOfTest(String conclusionsOfTest) {
        this.conclusionsOfTest = conclusionsOfTest;
    }

    public String getCorrectiveAction() {
        return correctiveAction;
    }

    public TestObject(String conclusionsOfTest, String correctiveAction) {
        this.conclusionsOfTest = conclusionsOfTest;
        this.correctiveAction = correctiveAction;
    }

    public void setCorrectiveAction(String correctiveAction) {
        this.correctiveAction = correctiveAction;

    }
}

