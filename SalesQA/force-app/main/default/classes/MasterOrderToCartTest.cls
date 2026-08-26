@IsTest
public with sharing class MasterOrderToCartTest {

    @TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
	
    }

    @IsTest
    public static void mainTest(){
    	MasterOrderToCart.provisoryCodeCoverageMethod();
    }

}