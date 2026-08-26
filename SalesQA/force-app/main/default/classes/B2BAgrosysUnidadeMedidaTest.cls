@IsTest
public class B2BAgrosysUnidadeMedidaTest {
    
	@isTest
    static void getUnidadeMedida200(){
        test.startTest();
        B2BAgrosysUnidadeMedida.FiltroRequisicao filtro = new B2BAgrosysUnidadeMedida.FiltroRequisicao();
        B2BAgrosysUnidadeMedida b2bUnidadeMedida = new B2BAgrosysUnidadeMedida();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_un":[{"med_codigo":"%","med_descricao":"PERCENTUAL","med_tipo":true},{"med_codigo":"0","med_descricao":"ZERO","med_tipo":true},{"med_codigo":"AM","med_descricao":"AMPOLA","med_tipo":true}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<UnidadeMedida__c> lstUnidadeBefore = [SELECT Id FROM UnidadeMedida__c];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bUnidadeMedida.getUnidadeMedida(filtro);
		List<UnidadeMedida__c> lstUnidadeAfter = [SELECT Id FROM UnidadeMedida__c];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstUnidadeBefore.size());
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(3, lstLogAfter.size());
        System.assertEquals(1, lstUnidadeAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getUnidadeMedida400(){
        test.startTest();
        B2BAgrosysUnidadeMedida.FiltroRequisicao filtro = new B2BAgrosysUnidadeMedida.FiltroRequisicao();
        B2BAgrosysUnidadeMedida b2bUnidadeMedida = new B2BAgrosysUnidadeMedida();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bUnidadeMedida.getUnidadeMedida(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}