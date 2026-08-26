@isTest
public class B2BCancelUnpaidOrdersSchedulableTest {
    @IsTest
    static void test(){
        
        Test.startTest();
        B2BCancelUnpaidOrdersSchedulable sc = new B2BCancelUnpaidOrdersSchedulable();
        sc.execute(null);
        Test.stopTest();
        
    }
}