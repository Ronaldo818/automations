@IsTest
public with sharing class CommerceVerbaflexHelperTest {

    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();

    }

    @IsTest
    public static void mainTest(){
	CommerceVerbaflexHelper.provisoryCodeCoverageMethod();
    }

}