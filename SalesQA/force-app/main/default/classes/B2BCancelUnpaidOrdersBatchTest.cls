@isTest
public with sharing class B2BCancelUnpaidOrdersBatchTest {
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
    static void testCancelOrer(){

        Account acc = [SELECT Id FROM Account LIMIT 1];
        Product2 p = [SELECT Id FROM Product2 LIMIT 1];

        //unpaid order
        Order o = new Order();
        o.AccountId = acc.Id;
        o.PaymentExpiration__c = Datetime.now();
        o.Status = 'Draft';
        o.Type = 'B2B Commerce';
        o.EffectiveDate = Date.today();
        o.MetodoPagamento__c = 'pix';
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
            Database.executeBatch(new B2BCancelUnpaidOrdersBatch(), 1);
        Test.stopTest();
        
    }

    @IsTest
    static void testNoOrdersToCancel(){

        Account acc = [SELECT Id FROM Account LIMIT 1];
        Product2 p = [SELECT Id FROM Product2 LIMIT 1];

        //paid order
        Order o1 = new Order();
        o1.AccountId = acc.Id;
        o1.PaymentExpiration__c = Datetime.now();
        o1.Status = 'Draft';
        o1.StatusPagamento__c = 'paid';
        o1.Type = 'B2B Commerce';
        o1.EffectiveDate = Date.today();
        o1.MetodoPagamento__c = 'pix';
        insert o1;

        OrderItem oi1 = new OrderItem();
        oi1.Product2Id = p.Id;
        oi1.Quantity = 1;
        oi1.ListPrice = 100.0;
        oi1.UnitPrice = 100.0;
        oi1.OrderId = o1.Id;
        insert oi1;

        o1.Status = 'Reserved';
        update o1;




        //order in time
        Order o2 = new Order();
        o2.AccountId = acc.Id;
        o2.PaymentExpiration__c = Datetime.now().addDays(1);
        o2.Status = 'Draft';
        o2.Type = 'B2B Commerce';
        o2.EffectiveDate = Date.today();
        o2.MetodoPagamento__c = 'pix';
        insert o2;


        OrderItem oi2 = new OrderItem();
        oi2.Product2Id = p.Id;
        oi2.Quantity = 1;
        oi2.ListPrice = 100.0;
        oi2.UnitPrice = 100.0;
        oi2.OrderId = o2.Id;
        insert oi2;

        o2.Status = 'Activated';
        update o2;
        
        Test.startTest();
            Database.executeBatch(new B2BCancelUnpaidOrdersBatch(), 1);
        Test.stopTest();
        
    }
}