@IsTest
public class B2BAgrosysFamiliaProdutoTest {
    
	@isTest
    static void getFamilia200(){
        test.startTest();
        B2BAgrosysFamiliaProduto.FiltroRequisicao filtro = new B2BAgrosysFamiliaProduto.FiltroRequisicao();
        B2BAgrosysFamiliaProduto b2bFamilia = new B2BAgrosysFamiliaProduto();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_familia":[{"fam_codigo":1,"fam_descricao":"RESFRIADO**"},{"fam_codigo":2,"fam_descricao":"CONGELADO**"},{"fam_codigo":3,"fam_descricao":"INTERRUPTORES"}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<FamiliaProduto__c> lstFamiliaBefore = [SELECT Id FROM FamiliaProduto__c];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bFamilia.getFamilia(filtro);
		List<FamiliaProduto__c> lstFamiliaAfter = [SELECT Id FROM FamiliaProduto__c];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstFamiliaBefore.size());
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(3, lstLogAfter.size());
        System.assertEquals(1, lstFamiliaAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getFamilia400(){
        test.startTest();
        B2BAgrosysFamiliaProduto.FiltroRequisicao filtro = new B2BAgrosysFamiliaProduto.FiltroRequisicao();
        B2BAgrosysFamiliaProduto b2bFamilia = new B2BAgrosysFamiliaProduto();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bFamilia.getFamilia(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}