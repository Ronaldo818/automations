@IsTest
public class B2BAgrosysCidadeTest {
    
	@isTest
    static void getCidade200(){
        test.startTest();
        B2BAgrosysCidade.FiltroRequisicao filtro = new B2BAgrosysCidade.FiltroRequisicao();
        B2BAgrosysCidade b2bCidade = new B2BAgrosysCidade();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_cidade":[{"cid_cidade":1,"cid_nome":"CARNEIRINHO","cid_ibge":3114550,"est_estado":"MG","pai_pais":1058},{"cid_cidade":2,"cid_nome":"AALBORG","cid_ibge":0,"est_estado":"AA","pai_pais":2321},{"cid_cidade":13,"cid_nome":"AMMAN","cid_ibge":0,"est_estado":"AA","pai_pais":4030}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<Cidade__c> lstCidadeBefore = [SELECT Id FROM Cidade__c];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bCidade.getCidade(filtro);
		List<Cidade__c> lstCidadeAfter = [SELECT Id FROM Cidade__c];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstCidadeBefore.size());
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(3, lstLogAfter.size());
        System.assertEquals(1, lstCidadeAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getCidade400(){
        test.startTest();
        B2BAgrosysCidade.FiltroRequisicao filtro = new B2BAgrosysCidade.FiltroRequisicao();
        B2BAgrosysCidade b2bCidade = new B2BAgrosysCidade();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bCidade.getCidade(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}