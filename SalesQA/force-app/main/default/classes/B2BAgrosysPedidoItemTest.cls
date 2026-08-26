@IsTest
public class B2BAgrosysPedidoItemTest {
	
    @TestSetup
    static void makeData() {
        
        Product2 newProduct = new Product2(
            Name = 'isTest',
        	PesoUnitario__c = 1,
        	StockKeepingUnit = '1',
        	CodigoProdutoERP__c = '1',
        	UnidadeMedida__c = 'KG');
		insert newProduct;
        
        Pricebook2 newPriceBook = new Pricebook2(Name = 'isTest');
		insert newPriceBook;
        
        PricebookEntry standardPrice = new PricebookEntry(
            Pricebook2Id = Test.getStandardPricebookId(),
            Product2Id = newProduct.Id,
            UnitPrice = 0,
            IsActive = true);
        insert standardPrice;
        
        PricebookEntry newPricebokEntry = new PricebookEntry(
			UnitPrice = 15,
    		Pricebook2Id = newPricebook.Id,
			Product2Id = newProduct.Id);
		insert newPricebokEntry;
        
        Account newAccount = new Account(
        	Name = 'isTest',
        	Unidade__c = 1,
        	Empresa__c = 1,
        	CodigoClienteERP__c = '1',
        	Ocorrencia__c = 1,
        	Lista__c = 27);
        insert newAccount;
        
        UnidadeMedida__c newUnidadeMedida = new UnidadeMedida__c(
            TipoMedida__c = false
        );
        insert newUnidadeMedida;
        
        UnidadeProduto__c newUnidadeProduto = new UnidadeProduto__c(
            Key__c = '1-1',
        	Unidade__c = 1,
        	Product__c = newProduct.Id,
        	UnidadePedido__c = newUnidadeMedida.Id
        );
        insert newUnidadeProduto;
        
        WebStore newStore = new WebStore(Name = 'isTest');
		insert newStore;
        
        WebCart newCart = new WebCart(
        	AccountId = newAccount.Id,
        	Name = 'isTest',
        	WebStoreId = newStore.Id,
        	Type = 'Cart',
        	Status = 'Active');
        insert newCart;
        
        CartDeliveryGroup newDeliveryGroup = new CartDeliveryGroup(
        	Name = 'isTest',
        	CartId = newCart.Id);
		insert newDeliveryGroup;	
       	
        CartItem newCartItem = new CartItem(
        	Name = 'isTest',
        	CartId = newCart.Id,
        	CartDeliveryGroupId = newDeliveryGroup.Id,
        	Product2Id = newProduct.Id,
        	Quantity = 1,
        	Type = 'Product',
        	SalesPrice = 2,
        	Sku = '1',
        	ListPrice = 3,
        	TotalPrice = 4);
        insert newCartItem;
    }
    
    @isTest
    static void calcularImposto200(){
        test.startTest();
        String accountId = [SELECT Id FROM Account LIMIT 1].Id;
        String cartId = [SELECT Id FROM WebCart WHERE Name = 'isTest' LIMIT 1].Id;
        B2BAgrosysRetorno b2bRetorno = new B2BAgrosysRetorno();
        B2BAgrosysPedidoItem b2bPedidoItem = new B2BAgrosysPedidoItem();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"ds":{"tp_pedido":[{"emp_empresa":1,"uni_unidade":1,"ped_codigo":"","ped_palm":"","cli_codigo":606,"ped_emissao":"2023-05-16","ped_entrega":null,"cpg_condpag":"","ped_cashback":0.00,"ped_observ":"","ped_obsnf":"","ped_encerrado":false,"ind_underline":true,"tp_pedidoitem":[{"emp_empresa":1,"uni_unidade":1,"ped_codigo":"","pro_codigo":104,"med_codigo":"KG","pit_qtpedi":1001.00,"pit_pspedi":2002.000,"pit_vlpreco":11.0549750250,"pit_seq":0,"pit_observa":"","pre_tiporeg":21,"pit_tabpre":1,"pit_status":0,"pit_ordcomp":"","pit_seqordcomp":0,"pit_vlprecoaux":11.0300000000,"pit_vlrst":0.00,"pit_vfcpst":0.00,"pit_vldesc":1361.36}]}]}}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2bPedidoItem.calcularImposto(accountId, cartId);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id, responseBody01__c FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
        System.assertEquals(1, lstLogAfter.size());
		System.assertEquals(true, b2bRetorno != null);
        test.stopTest();
    }
    
    @isTest
    static void calcularImpostoError(){
        String accountId = [SELECT Id FROM Account LIMIT 1].Id;
        String cartId = [SELECT Id FROM WebCart WHERE Name = 'isTest' LIMIT 1].Id;
        B2BAgrosysRetorno b2bRetorno = new B2BAgrosysRetorno();
        B2BAgrosysPedidoItem b2bPedidoItem = new B2BAgrosysPedidoItem();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        test.startTest();
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2bPedidoItem.calcularImposto(accountId, cartId);
        List<IntegrationLog__c> lstLogAfter = [SELECT Id, responseBody01__c FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
        System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}