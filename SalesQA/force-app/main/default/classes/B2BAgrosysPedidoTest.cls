@IsTest
public class B2BAgrosysPedidoTest {

	@TestSetup
    static void makeData() {
        
        Product2 newProduct = new Product2(
            Name = 'isTest',
        	PesoUnitario__c = 6,
        	StockKeepingUnit = '1',
        	CodigoProdutoERP__c = '1');
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
        
        CondicaoPagamento__c newCondPagto = new CondicaoPagamento__c(
        	AVista__c = true,
        	CodigoCondicaoPagamento__c = '1',
        	Descricao__c = 'isTest',
        	Name = 'isTest',
        	NumeroParcelas__c = 1);
        insert newCondPagto;
        
        Account newAccount = new Account(
        	Name = 'isTest',
        	Unidade__c = 1,
        	Empresa__c = 1,
        	CodigoClienteERP__c = '1',
        	CondicaoPagamento__c = newCondPagto.Id,
        	Ocorrencia__c = 1,
        	Lista__c = 27);
        insert newAccount;
        
        Order newOrder = new Order(
        	EffectiveDate = Date.today(),
        	Status = 'Draft',
        	ValorCashback__c = 2,
        	AccountId = newAccount.Id,
        	Pricebook2Id = newPriceBook.Id,
            NumeroPedidoAgrosys__c = '1000'
        );
        insert newOrder;
        
        OrderItem newOrderItem = new OrderItem(
        	OrderId = newOrder.Id,
        	PricebookEntryId = newPricebokEntry.Id,
        	Product2Id = newProduct.Id,
        	ValorDesconto__c = 0,
        	ValorST__c = 0,
        	ListPrice = 15,
        	FundoCombatePobrezaST__c = 0,
        	LineNumber = 1,
        	UnitPrice = 15,
        	Quantity = 1);
        insert newOrderItem;
        
        UnidadeProduto__c UnidadeProduto = new UnidadeProduto__c(
            Key__c = '1-1',
        	Unidade__c = 1);
        insert UnidadeProduto;
        
        Order orderError = new Order(
        	NumeroPedidoAgrosys__c = '1111',
        	EffectiveDate = Date.today(),
        	Status = 'Draft',
        	ValorCashback__c = 2,
        	AccountId = newAccount.Id,
        	Pricebook2Id = newPriceBook.Id);
        insert orderError;
    }
    
    @isTest
    static void setPedidoSuccess01(){
        test.startTest();

		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" :1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
        B2BAgrosysPedido.SendOrderParameter parameter = new B2BAgrosysPedido.SendOrderParameter();
        parameter.orderId = [SELECT Id FROM Order LIMIT 1].Id;
        parameter.parcelas = '1';
        parameter.formaPagamento = 'boleto';
		parameter.operacao = 2;

		String retorno = B2BAgrosysPedido.setPedido(System.JSON.serialize(parameter));
        
        System.assertEquals(0, lstLogBefore.size());
		
        test.stopTest();

		List<IntegrationLog__c> lstLogAfter = [SELECT Id, responseBody01__c FROM IntegrationLog__c];
        System.assertEquals(1, lstLogAfter.size());
    }
    
    @isTest
    static void setPedidoSuccess02(){
        test.startTest();

		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" :2,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		B2BAgrosysPedido.SendOrderParameter parameter = new B2BAgrosysPedido.SendOrderParameter();
        parameter.orderId = [SELECT Id FROM Order LIMIT 1].Id;
        parameter.formaPagamento = 'pix';
		parameter.operacao = 2;

		String retorno = B2BAgrosysPedido.setPedido(System.JSON.serialize(parameter));
        
        System.assertEquals(0, lstLogBefore.size());
		
        test.stopTest();

		List<IntegrationLog__c> lstLogAfter = [SELECT Id, responseBody01__c FROM IntegrationLog__c];
        System.assertEquals(1, lstLogAfter.size());
    }
    
    @isTest
    static void setPedidoSuccess03(){
        test.startTest();

		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" :1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		B2BAgrosysPedido.SendOrderParameter parameter = new B2BAgrosysPedido.SendOrderParameter();
        parameter.orderId = [SELECT Id FROM Order LIMIT 1].Id;
        parameter.parcelas = '1';
        parameter.formaPagamento = 'cartao de credito';
		parameter.operacao = 2;
        parameter.cvvCartao = '444';
        parameter.anoCartao = '28';
        parameter.mesCartao = '05';
        parameter.numeroCartao = '4444444444444444';
        parameter.nomeCartao = 'isTest';

		String retorno = B2BAgrosysPedido.setPedido(System.JSON.serialize(parameter));
        
        System.assertEquals(0, lstLogBefore.size());
		
        test.stopTest();

		List<IntegrationLog__c> lstLogAfter = [SELECT Id, responseBody01__c FROM IntegrationLog__c];
        System.assertEquals(2, lstLogAfter.size());
    }
    
    @isTest
    static void setPedidoSuccess04(){
        test.startTest();

		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" :797,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		B2BAgrosysPedido.SendOrderParameter parameter = new B2BAgrosysPedido.SendOrderParameter();
        parameter.orderId = [SELECT Id FROM Order LIMIT 1].Id;
        parameter.parcelas = '1';
        parameter.formaPagamento = 'boleto';
		parameter.operacao = 2;

		String retorno = B2BAgrosysPedido.setPedido(System.JSON.serialize(parameter));
        
        System.assertEquals(0, lstLogBefore.size());
		
        test.stopTest();

		List<IntegrationLog__c> lstLogAfter = [SELECT Id, responseBody01__c FROM IntegrationLog__c];
        System.assertEquals(1, lstLogAfter.size());
    }
    
    @isTest
    static void setPedidoError01(){
        test.startTest();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"status":1,"msg":"Recebimentook.","obj":{"vis_codigo":1111111111,"mensintegra":""}}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
        B2BAgrosysPedido.SendOrderParameter parameter = new B2BAgrosysPedido.SendOrderParameter();
        parameter.orderId = [SELECT Id FROM Order LIMIT 1].Id;
        parameter.parcelas = '1';
        parameter.formaPagamento = 'boleto';
		parameter.operacao = 2;

		String retorno = B2BAgrosysPedido.setPedido(System.JSON.serialize(parameter));
        
        System.assertEquals(0, lstLogBefore.size());
		
        test.stopTest();
    }
    
     @isTest
    static void setPedidoError02(){
        test.startTest();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"status":1,"msg":"Recebimentook.","obj":{"vis_codigo":1111111111,"mensintegra":""}}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
        B2BAgrosysPedido.SendOrderParameter parameter = new B2BAgrosysPedido.SendOrderParameter();
        parameter.orderId = [SELECT Id FROM Order LIMIT 1].Id;
        parameter.formaPagamento = 'error setrequisicao';

		String retorno = B2BAgrosysPedido.setPedido(System.JSON.serialize(parameter));
        
        System.assertEquals(0, lstLogBefore.size());
		
        test.stopTest();
    }
    
    @isTest
    static void setPedidoError03(){
        test.startTest();

		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" :1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
        B2BAgrosysPedido.SendOrderParameter parameter = new B2BAgrosysPedido.SendOrderParameter();
        parameter.orderId = [SELECT Id FROM Order LIMIT 1].Id;
        parameter.parcelas = '1';
        parameter.formaPagamento = 'error setpedido';
		parameter.operacao = 2;

		String retorno = B2BAgrosysPedido.setPedido(System.JSON.serialize(parameter));
        
        System.assertEquals(0, lstLogBefore.size());
		
        test.stopTest();

		List<IntegrationLog__c> lstLogAfter = [SELECT Id, responseBody01__c FROM IntegrationLog__c];
        System.assertEquals(2, lstLogAfter.size());
    }
    
    @isTest
    static void setPedidoError04(){
        test.startTest();

		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo" :1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
        B2BAgrosysPedido.SendOrderParameter parameter = new B2BAgrosysPedido.SendOrderParameter();
        parameter.orderId = [SELECT Id FROM Order LIMIT 1].Id;
        parameter.parcelas = 'error setpedido';
        parameter.formaPagamento = 'pix';
		parameter.operacao = 2;

		String retorno = B2BAgrosysPedido.setPedido(System.JSON.serialize(parameter));
        
        System.assertEquals(0, lstLogBefore.size());
		
        test.stopTest();

		List<IntegrationLog__c> lstLogAfter = [SELECT Id, responseBody01__c FROM IntegrationLog__c];
        System.assertEquals(2, lstLogAfter.size());
    }
}