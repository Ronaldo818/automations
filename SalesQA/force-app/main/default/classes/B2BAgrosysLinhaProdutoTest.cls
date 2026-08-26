@IsTest
public class B2BAgrosysLinhaProdutoTest {
    
	@isTest
    static void getLinha200(){
        test.startTest();
        B2BAgrosysLinhaProduto.FiltroRequisicao filtro = new B2BAgrosysLinhaProduto.FiltroRequisicao();
        B2BAgrosysLinhaProduto b2bLinha = new B2BAgrosysLinhaProduto();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_linha":[{"lin_codigo":1,"lin_descricao":"OVOCOMERCIAL"},{"lin_codigo":2,"lin_descricao":"OVOFERTIL"},{"lin_codigo":3,"lin_descricao":"PINTODE1DIA"}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<LinhaProduto__c> lstLinhaBefore = [SELECT Id FROM LinhaProduto__c];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bLinha.getLinha(filtro);
		List<LinhaProduto__c> lstLinhaAfter = [SELECT Id FROM LinhaProduto__c];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLinhaBefore.size());
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(3, lstLogAfter.size());
        System.assertEquals(1, lstLinhaAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getLinha400(){
        test.startTest();
        B2BAgrosysLinhaProduto.FiltroRequisicao filtro = new B2BAgrosysLinhaProduto.FiltroRequisicao();
        B2BAgrosysLinhaProduto b2bLinha = new B2BAgrosysLinhaProduto();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bLinha.getLinha(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}