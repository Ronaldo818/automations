@IsTest
public class B2BAgrosysGrupoProdutoTest {
    
	@isTest
    static void getGrupoItem200(){
        test.startTest();
        B2BAgrosysGrupoProduto.FiltroRequisicao filtro = new B2BAgrosysGrupoProduto.FiltroRequisicao();
        B2BAgrosysGrupoProduto b2bLinha = new B2BAgrosysGrupoProduto();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_grpitem":[{"gru_codigo":1,"gru_descricao":"EmbalagensSecundarias"},{"gru_codigo":2,"gru_descricao":"Materiaisp/Enfermaria-Medicamentos"},{"gru_codigo":3,"gru_descricao":"Vacinas"}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<GrupoProduto__c> lstGrupoBefore = [SELECT Id FROM GrupoProduto__c];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bLinha.getGrupoItem(filtro);
		List<GrupoProduto__c> lstGrupoAfter = [SELECT Id FROM GrupoProduto__c];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstGrupoBefore.size());
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(3, lstLogAfter.size());
        System.assertEquals(1, lstGrupoAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getGrupoItem400(){
        test.startTest();
        B2BAgrosysGrupoProduto.FiltroRequisicao filtro = new B2BAgrosysGrupoProduto.FiltroRequisicao();
        B2BAgrosysGrupoProduto b2bLinha = new B2BAgrosysGrupoProduto();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bLinha.getGrupoItem(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}