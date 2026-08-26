@ISTest
public class B2BAgrosysTabelaPrecoTest  {
    
    @TestSetup
    static void makeData() {
        Account newAcc = new Account(
            Name				= 'isTest',
			CodigoClienteERP__c	= '999',
			Empresa__c 			= 1,
			unidade__c 			= 1
        );
        insert newAcc;
    }
    
	@isTest
    static void getTabela200(){
        test.startTest();
        B2BAgrosysTabelaPreco b2bGetTabela = new B2BAgrosysTabelaPreco();
        B2BAgrosysRetorno b2bRetorno = new B2BAgrosysRetorno();

		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_tabpreco":[{"emp_empresa":1,"uni_unidade":1,"pre_tiporeg":21,"pit_tabpre":1,"pro_codigo":1039,"pre_dtatual":"2023-06-27","pre_tabela":11.8800,"pre_vlrst":0.1210683333,"pre_vfcpst":0.0000000000},{"emp_empresa":1,"uni_unidade":1,"pre_tiporeg":22,"pit_tabpre":1,"pro_codigo":1039,"pre_dtatual":"2023-06-27","pre_tabela":11.9500,"pre_vlrst":0.1218033333,"pre_vfcpst":0.0000000000},{"emp_empresa":1,"uni_unidade":1,"pre_tiporeg":23,"pit_tabpre":1,"pro_codigo":1039,"pre_dtatual":"2023-06-27","pre_tabela":12.0100,"pre_vlrst":0.1224333333,"pre_vfcpst":0.0000000000}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		Account acc = [
            SELECT Name, CodigoClienteERP__c, Empresa__c, unidade__c
            FROM Account
            LIMIT 1
        ];
        List<Pricebook2> lstTabelaBefore = [SELECT Id FROM Pricebook2];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2bGetTabela.getCustomerPricebook2(acc);
		List<Pricebook2> lstTabelaAfter = [SELECT Id FROM Pricebook2];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0,lstTabelaBefore.size());
        System.assertEquals(0,lstLogBefore.size());
        System.assertEquals(0, lstTabelaAfter.size());
		System.assertEquals(1,lstLogAfter.size());
		System.assertEquals(true,b2bRetorno != null);
        test.stopTest();
    }
    
    @isTest
    static void getTabela400(){
        test.startTest();
        B2BAgrosysRetorno b2bRetorno = new B2BAgrosysRetorno();
        B2BAgrosysTabelaPreco b2bGetTabela = new B2BAgrosysTabelaPreco();
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        Account acc = [
            SELECT Name, CodigoClienteERP__c, Empresa__c, unidade__c
            FROM Account
            LIMIT 1
        ];
        List<Pricebook2> lstTabelaBefore = [SELECT Id FROM Pricebook2];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2bGetTabela.getCustomerPricebook2(acc);
        List<Pricebook2> lstTabelaAfter = [SELECT Id FROM Pricebook2];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0,lstTabelaBefore.size());
        System.assertEquals(0,lstLogBefore.size());
        System.assertEquals(0, lstTabelaAfter.size());
		System.assertEquals(1,lstLogAfter.size());
		System.assertEquals(true,b2bRetorno != null);
        test.stopTest();
    }
}