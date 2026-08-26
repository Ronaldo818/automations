@isTest
private class CustomOrderListViewControllerTest {

    @TestSetup
    static void makeData() {
        Account newAccount = new Account(
        	Name = 'isTest',
        	CodigoCliente__c = 'isTest01');
        insert newAccount;
	
        
        CondicaoPagamento__c newCondPagamento = new CondicaoPagamento__c(
       		Name = 'isTest',
        	CodigoCondicaoPagamento__c = '1');
        insert newCondPagamento;
       	
        Contact newContact = new Contact(
            LastName ='testCon',
            AccountId = newAccount.Id);
        insert newContact;
        
    	WebStore newStore = new WebStore(
            Name = 'isTest');
		insert newStore;

        WebCart newCart = new WebCart(
        	AccountId 				= newAccount.Id,
        	Name 					= 'isTest',
        	WebStoreId 				= newStore.Id,
        	Type 					= 'Cart',
        	Status 					= 'Active'
        );
        insert newCart;

        CartDeliveryGroup newDeliveryGroup = new CartDeliveryGroup(
            Name 	= 'Default Delivery',
            CartId	= newCart.Id
        );
        insert newDeliveryGroup;

        Product2 parentProduct = new Product2(
            Name = 'isTest01',
            IsActive = true,
        	StockKeepingUnit = '-0001',
            ProductCode = '-0001');
		insert parentProduct;

        Product2 childProduct = new Product2(
            Name = 'isTest02',
            IsActive = true,
        	StockKeepingUnit = '-0001_BRA',
            ProductCode = '-0001_BRA');
		insert childProduct;
        
        Pricebook2 newPriceBook = new Pricebook2(Name = '-TUV18002205 A', IsActive = true);
		insert newPriceBook;

        PricebookEntry newStandardPricebokEntry01 = new PricebookEntry(
			UnitPrice = 0.0,
    		Pricebook2Id = Test.getStandardPricebookId(),
			Product2Id = parentProduct.Id
        );
		insert newStandardPricebokEntry01;
                
        PricebookEntry newPricebokEntry01 = new PricebookEntry(
			UnitPrice = 15.0,
    		Pricebook2Id = newPriceBook.Id,
			Product2Id = parentProduct.Id
            );
		insert newPricebokEntry01;

        Transportadora__c trans = new Transportadora__c();
        trans.CodigoLegado__c = '9872000174';
        insert trans;

        Order o  = new Order();
        o.AccountId = newAccount.Id;
        o.Status = 'Draft';
        o.EffectiveDate = Date.today();
        insert o;

        OrderItem oiP = new OrderItem();
        oiP.CodigoItemPedido__c = '001_490111';
        oiP.OrderId = o.Id;
        oiP.Product2Id = parentProduct.Id;
        oiP.SequenciaItem__c = 2;
        oiP.Quantity = 96.00;
        oiP.UnitPrice = 15.988800;
        oiP.ListPrice = 15.988800;

        insert oiP;

        OrderItem oiC = new OrderItem();
        oiC.CodigoItemPedido__c = 'BRA_490111';
        oiC.OrderId = o.Id;
        oiC.Product2Id = childProduct.Id;
        oiC.SequenciaItem__c = 2;
        oiC.Quantity = 96.00;
        oiC.UnitPrice = 15.988800;
        oiC.ListPrice = 15.988800;
        insert oiC;
    }

    @isTest
    static void test(){

        Test.startTest();

           CustomOrderListViewController.getObjectRecords('None', 'OrderNumber,CodigoPedidoERP__c,CreatedDate', 'Order','CreatedDate', 'ASC', '0');
           CustomOrderListViewController.getObjectRecords('Id != null', 'OrderNumber,CodigoPedidoERP__c,CreatedDate', 'Order','CreatedDate', 'DESC', '0');
           CustomOrderListViewController.getObjectRecords('Cliente', 'OrderNumber,CodigoPedidoERP__c,CreatedDate', 'Order','CreatedDate', 'DESC', '0');
           CustomOrderListViewController.buscar('490111', 'OrderNumber,CodigoPedidoERP__c,CreatedDate');
        Test.stopTest();
    }


}