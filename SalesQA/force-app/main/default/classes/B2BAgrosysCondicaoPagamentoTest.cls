@IsTest
public class B2BAgrosysCondicaoPagamentoTest {
    
	@isTest
    static void getCondPagto200(){
        test.startTest();
        B2BAgrosysCondicaoPagamento.FiltroRequisicao filtro = new B2BAgrosysCondicaoPagamento.FiltroRequisicao();
        B2BAgrosysCondicaoPagamento b2bCondicaoPagamento = new B2BAgrosysCondicaoPagamento();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_condpagto":[{"cpg_condpag":"02BC","cpg_descricao":"02DIASBANCO","cpg_avista":false,"cpg_nrparcelas":1},{"cpg_condpag":"02CA","cpg_descricao":"02DIASCARTEIRA","cpg_avista":false,"cpg_nrparcelas":1},{"cpg_condpag":"02X7","cpg_descricao":"07/14DIASBANCO","cpg_avista":false,"cpg_nrparcelas":2}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<CondicaoPagamento__c> lstCondicaoBefore = [SELECT Id FROM CondicaoPagamento__c];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bCondicaoPagamento.getCondPagto(filtro);
		List<CondicaoPagamento__c> lstCondicaoAfter = [SELECT Id FROM CondicaoPagamento__c];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstCondicaoBefore.size());
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(3, lstLogAfter.size());
        System.assertEquals(1, lstCondicaoAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getCondPagto400(){
        test.startTest();
        B2BAgrosysCondicaoPagamento.FiltroRequisicao filtro = new B2BAgrosysCondicaoPagamento.FiltroRequisicao();
        B2BAgrosysCondicaoPagamento b2bCondicaoPagamento = new B2BAgrosysCondicaoPagamento();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bCondicaoPagamento.getCondPagto(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}