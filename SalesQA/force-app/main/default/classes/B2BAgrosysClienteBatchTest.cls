@IsTest
public class B2BAgrosysClienteBatchTest {
	
    @TestSetup
    static void makeData() {
    	List<Account> lstAccountUpsert = new List<Account>();
        for(Integer i = 0; i < 3; i++) {
        	Account newAcc = new Account(
                Name				= 'isTest ' + i,
                CodigoClienteERP__c	= String.valueOf(i),
				Empresa__c 			= 1,
				unidade__c 			= 1);
            lstAccountUpsert.add(newAcc);
        }
        upsert lstAccountUpsert;    
    }
    
    @isTest
    static void b2bAgrosysClienteBatchSuccess01() {
        test.startTest();
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_cliente":[{"cli_codigo":83358,"emp_empresa":1,"uni_unidade":1,"cli_fantasi":"LIDIANEGONCALVESDEOLIVEIRA10086401661","cli_razsoc":"LIDIANEGONCALVESDEOLIVEIRA10086401661","cli_cnpj":"47747783000135","cli_aplicast":true,"cli_email":"test01@test.com;test02@test.com","cli_foncom":"37999213344","cli_dtativ":"2022-10-18","cli_dtinativ":"2023-07-10","cli_overprice":3.0,"cli_limcre":2600.0,"cpg_condpag":"07BC","fxp_tpcli":1,"pre_tiporeg":27,"pit_tabpre":1,"cli_ecommerce":true}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		B2BAgrosysClienteBatch b2bBatch = new B2BAgrosysClienteBatch();
        Database.executeBatch(b2bBatch);
		
        System.assertEquals(0, lstLogBefore.size());
        test.stopTest();
        List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        System.assertEquals(2, lstLogAfter.size());
    }
    
    @isTest
    static void b2bAgrosysClienteBatchSuccess02() {
        test.startTest();
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_cliente":[{"cli_codigo":83358,"emp_empresa":1,"uni_unidade":1,"cli_fantasi":"LIDIANEGONCALVESDEOLIVEIRA10086401661","cli_razsoc":"LIDIANEGONCALVESDEOLIVEIRA10086401661","cli_cnpj":"47747783000135","cli_aplicast":true,"cli_email":"test01@test.com;test02@test.com","cli_foncom":"37999213344","cli_dtativ":"2022-10-18","cli_dtinativ":"2023-07-10","cli_overprice":3.0,"cli_limcre":2600.0,"cpg_condpag":"07BC","fxp_tpcli":1,"pre_tiporeg":27,"pit_tabpre":1,"cli_ecommerce":true}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		B2BAgrosysClienteBatch b2bBatch = new B2BAgrosysClienteBatch(1);
        Database.executeBatch(b2bBatch);
		
        System.assertEquals(0, lstLogBefore.size());
        test.stopTest();
        List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        System.assertEquals(2, lstLogAfter.size());
    }

    @isTest
    static void b2bAgrosysClienteBatchError02() {
        test.startTest();
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('Erro de conexao: Retorno Invalido');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		B2BAgrosysClienteBatch b2bBatch = new B2BAgrosysClienteBatch(1);
        Database.executeBatch(b2bBatch);
		
        System.assertEquals(0, lstLogBefore.size());
        test.stopTest();
        List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        System.assertEquals(2, lstLogAfter.size());

        List<LogErro__c> lstErrorLogs = [SELECT Id FROM LogErro__c];
        System.assertEquals(1, lstErrorLogs.size());
    }
    
}