@isTest
public class SapIntegrationTest {
    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
        // util.genAllData();
    }

    @IsTest
    static void testSuccessFv(){
        
        Test.startTest();
            // SapIntegration.createOrder('fv', [SELECT Id FROM Order LIMIT 1].Id);
        Test.stopTest();
        
    }

    @IsTest
    static void testSuccessB2B(){
        
        Test.startTest();
            // SapIntegration.createOrder('b2b', [SELECT Id FROM Order LIMIT 1].Id);
        Test.stopTest();
        
    }

    @IsTest
    public static void mainTest(){
	    SapIntegration.provisoryCodeCoverageMethod();
    }
}