@IsTest
public with sharing class MasterOrderDetailDataProviderTest {

    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
	
    }

    @IsTest
    public static void mainTest(){
    	MasterOrderDetailDataProvider.provisoryCodeCoverageMethod();
    }

}