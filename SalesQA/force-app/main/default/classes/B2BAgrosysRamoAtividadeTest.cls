@IsTest
public class B2BAgrosysRamoAtividadeTest {
    
	@isTest
    static void getRamoAtivade200(){
        test.startTest();
        B2BAgrosysRamoAtividade.FiltroRequisicao filtro = new B2BAgrosysRamoAtividade.FiltroRequisicao();
        B2BAgrosysRamoAtividade b2bRamo = new B2BAgrosysRamoAtividade();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integraçãorealizadacomsucesso","msg_retorno":[{"tp_rm":[{"ram_codigo":1,"ram_nome":"AVICULTURA","ram_sigla":""},{"ram_codigo":2,"ram_nome":"SUPERMERCADO","ram_sigla":""},{"ram_codigo":3,"ram_nome":"DISTRIBUIDORAS","ram_sigla":""}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<RamoAtividade__c> lstRamoBefore = [SELECT Id FROM RamoAtividade__c];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRamo.getRamoAtivade(filtro);
		List<RamoAtividade__c> lstRamoAfter = [SELECT Id FROM RamoAtividade__c];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstRamoBefore.size());
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(3, lstLogAfter.size());
        System.assertEquals(1, lstRamoAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getRamoAtivade400(){
        test.startTest();
        B2BAgrosysRamoAtividade.FiltroRequisicao filtro = new B2BAgrosysRamoAtividade.FiltroRequisicao();
        B2BAgrosysRamoAtividade b2bRamo = new B2BAgrosysRamoAtividade();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRamo.getRamoAtivade(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}