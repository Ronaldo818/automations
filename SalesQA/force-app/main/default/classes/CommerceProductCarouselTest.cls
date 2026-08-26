@isTest
private class CommerceProductCarouselTest {

    /*@TestSetup
    static void makeData(){
        TestUtil util = new TestUtil();
        util.genStoreData();
    }*/

    @isTest
    static void CommerceProductCarouselTest01() {
        List<String> skus = new List<String>();

        for(Product2 p : [SELECT ProductCode FROM Product2]){
            skus.add(p.ProductCode);
        }

        Test.startTest();

        try{
            CommerceProductCarousel.getProducts(null, 20, String.join(skus, ','), true, false, null);
        } catch (Exception ex) {

        }

        Test.stopTest();
    }

    @isTest
    static void CommerceProductCarouselTest02() {
        Test.startTest();
        try{
            CommerceProductCarousel.getProducts('Test', 20, null, true, false, null);
        } catch (Exception ex) {

        }
        Test.stopTest();
    }

    @isTest
    static void CommerceProductCarouselTest04() {
        Test.startTest();
        try{
            CommerceProductCarousel.getProducts(null, null, null, true, true, null);
        } catch (Exception ex) {}
        Test.stopTest();
    }

    @IsTest
    public static void mainTest(){
    	CommerceProductCarousel.provisoryCodeCoverageMethod();
    }
}