@IsTest
public class B2BAgrosysCashbackTest {
    @testSetup
    static void makeData() {
    	Account newAccount = new Account(
        	Name = 'isTest',
        	Unidade__c = 1,
        	Empresa__c = 1,
        	CodigoClienteERP__c = '1',
        	Ocorrencia__c = 1,
        	Lista__c = 27);
        insert newAccount;
        
        UnidadeProduto__c UnidadeProduto = new UnidadeProduto__c(
            Key__c = '1-1',
        	Unidade__c = 1);
        insert UnidadeProduto;
        
        WebStore newStore = new WebStore(Name = 'isTest');
		insert newStore;
        
        WebCart newCart = new WebCart(
        	AccountId = newAccount.Id,
        	Name = 'isTest',
        	WebStoreId = newStore.Id,
        	Type = 'Cart',
        	Status = 'Active');
        insert newCart;    
    }
    
	@isTest
    static void getCashbackSuccess01(){
        test.startTest();
        B2BAgrosysCashback.CashbackReturn b2bRetorno = new B2BAgrosysCashback.CashbackReturn();
        B2BAgrosysCashback b2bCashback = new B2BAgrosysCashback();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"","msg_retorno":[{"tp_extverba":[{"cli_codigo":606,"hvr_dtmov":"2022-10-20","hvr_tpverba":0,"hvr_tpmov":false,"hvr_tpdoc":"SaldodoDia","hvr_docto":"","hvr_valor":0.0,"ite_codigo":0,"hvr_saldo":500.0,"hvr_seq":8}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2bCashback.getCashback('606');
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(true, b2bRetorno != null);
        System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getCashbackSuccess02(){
        test.startTest();
        B2BAgrosysCashback.CashbackReturn b2bRetorno = new B2BAgrosysCashback.CashbackReturn();
        B2BAgrosysCashback b2bCashback = new B2BAgrosysCashback();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		// mock.setResponseBody('');
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":2,"msg_descricao":"","msg_retorno":[{"tp_extverba":[{"cli_codigo":606,"hvr_dtmov":"2022-10-20","hvr_tpverba":0,"hvr_tpmov":false,"hvr_tpdoc":"SaldodoDia","hvr_docto":"","hvr_valor":0.0,"ite_codigo":0,"hvr_saldo":500.0,"hvr_seq":8}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2bCashback.getCashback('606');
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(false, b2bRetorno == null);
        System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}