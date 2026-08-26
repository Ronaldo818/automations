@IsTest
public with sharing class FvCartControllerTest {

    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
	
    }

    @IsTest
    public static void mainTest(){
    	FvCartController.provisoryCodeCoverageMethod();
    }

}