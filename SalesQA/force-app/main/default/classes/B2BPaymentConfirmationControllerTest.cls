@isTest
public class B2BPaymentConfirmationControllerTest {
    @TestSetup
    static void makeData(){
        Account acc = new Account();
        acc.Name = 'Test';
        insert acc;

        Product2 p = new Product2();
        p.Name = 'Test Product';
        p.ProductCode = '1234';
        p.StockKeepingUnit = '1234';
        insert p;
    }

    @IsTest
    static void test(){
        Account acc = [SELECT Id FROM Account LIMIT 1];
        Product2 p = [SELECT Id FROM Product2 LIMIT 1];

        //unpaid order
        Order o = new Order();
        o.AccountId = acc.Id;
        o.PaymentExpiration__c = Datetime.now();
        o.Status = 'Draft';
        o.Type = 'B2B Commerce';
        o.EffectiveDate = Date.today();
        insert o;

        OrderItem oi = new OrderItem();
        oi.Product2Id = p.Id;
        oi.Quantity = 1;
        oi.ListPrice = 100.0;
        oi.UnitPrice = 100.0;
        oi.OrderId = o.Id;
        insert oi;

        o.Status = 'Activated';
        update o;
        Test.startTest();
            B2BPaymentConfirmationController.getPaymentData([SELECT OrderReferenceNumber FROM Order WHERE id = :o.Id LIMIT 1][0].OrderReferenceNumber);
        Test.stopTest();
        
    }
}