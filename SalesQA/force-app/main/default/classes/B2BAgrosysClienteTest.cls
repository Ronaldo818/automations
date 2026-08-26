@IsTest
public class B2BAgrosysClienteTest {
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
    static void getCliente200(){
        test.startTest();
        List<B2BAgrosysRetorno.TpCliente> b2bRetorno = new List<B2BAgrosysRetorno.TpCliente>();
        B2BAgrosysCliente.FiltroRequisicao filtro = new B2BAgrosysCliente.FiltroRequisicao();
        B2BAgrosysCliente b2BCliente = new B2BAgrosysCliente();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_cliente":[{"cli_codigo":256,"emp_empresa":1,"uni_unidade":1,"cli_fantasi":"CELIOALVESXAVIER","cli_razsoc":"CELIOALVESXAVIER","cli_email":"test01@test.com","cli_fisjur":true,"cli_cpf":"71417826649","cli_foncom":"3732149593","cli_fonfin":"3799630663","cli_dtativ":"2006-02-08","cli_dtinativ":"2023-07-10","cli_overprice":3.0,"cli_limcre":1000.0,"cpg_condpag":"07BC","fxp_tpcli":1,"pre_tiporeg":27,"pit_tabpre":1,"cli_ecommerce":true}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2BCliente.getCliente(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
        System.assertEquals(true, b2bRetorno != null);
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getCliente400(){
        test.startTest();
        List<B2BAgrosysRetorno.TpCliente> b2bRetorno = new List<B2BAgrosysRetorno.TpCliente>();
        B2BAgrosysCliente.FiltroRequisicao filtro = new B2BAgrosysCliente.FiltroRequisicao();
        B2BAgrosysCliente b2BCliente = new B2BAgrosysCliente();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2BCliente.getCliente(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
        System.assertEquals(true, b2bRetorno == null);
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }

    @isTest
    static void clienteBatch200() {
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
    static void clienteSchedule200() {
        test.startTest();
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_cliente":[{"cli_codigo":83358,"emp_empresa":1,"uni_unidade":1,"cli_fantasi":"LIDIANEGONCALVESDEOLIVEIRA10086401661","cli_razsoc":"LIDIANEGONCALVESDEOLIVEIRA10086401661","cli_cnpj":"47747783000135","cli_aplicast":true,"cli_email":"test01@test.com;test02@test.com","cli_foncom":"37999213344","cli_dtativ":"2022-10-18","cli_dtinativ":"2023-07-10","cli_overprice":3.0,"cli_limcre":2600.0,"cpg_condpag":"07BC","fxp_tpcli":1,"pre_tiporeg":27,"pit_tabpre":1,"cli_ecommerce":true}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
        
		Datetime newSchuledatetime = System.now().addMinutes(1);
        String sch = '0 ' + String.valueof(newSchuledatetime.minute())+ ' ' + String.valueof(newSchuledatetime.hour()) + ' ' + String.valueof(newSchuledatetime.day()) + ' ' + String.valueof(newSchuledatetime.month()) + ' ? ' + String.valueof(newSchuledatetime.year());
        B2BScheduleAgrosys rotina = new B2BScheduleAgrosys();
		System.schedule('Class Test B2BScheduleAgrosys - ' + newSchuledatetime, sch, rotina);
		
        System.assertEquals(0, lstLogBefore.size());
        test.stopTest();
        
        List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        System.assertEquals(0, lstLogAfter.size());
    }
}