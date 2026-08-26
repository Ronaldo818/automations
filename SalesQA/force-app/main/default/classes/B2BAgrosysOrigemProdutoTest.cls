@IsTest
global class B2BAgrosysOrigemProdutoTest {
    
	@isTest
    static void getOrigem200(){
        test.startTest();
        B2BAgrosysOrigemProduto.FiltroRequisicao filtro = new B2BAgrosysOrigemProduto.FiltroRequisicao();
        B2BAgrosysOrigemProduto b2bGetOrigem = new B2BAgrosysOrigemProduto();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_origem":[{"ori_codigo":1,"ori_descricao":"FRIGORIFICO"},{"ori_codigo":2,"ori_descricao":"ANIMALVIVO"},{"ori_codigo":3,"ori_descricao":"ITEMDESUPRIMENTOS"},{"ori_codigo":4,"ori_descricao":"SUB-PRODUTO"},{"ori_codigo":5,"ori_descricao":"PINTODE1DIA"},{"ori_codigo":6,"ori_descricao":"BI-BUSSINESSINTELIGENCE"},{"ori_codigo":7,"ori_descricao":"RAÇAO"},{"ori_codigo":8,"ori_descricao":"OVO"},{"ori_codigo":9,"ori_descricao":"INDUSTRIALIZADOS"},{"ori_codigo":10,"ori_descricao":"REVENDA"},{"ori_codigo":11,"ori_descricao":"MARCAPROPRIA"}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<OrigemProduto__c> lstOrigemBefore = [SELECT Id FROM OrigemProduto__c];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bGetOrigem.getOrigem(filtro);
		List<OrigemProduto__c> lstOrigemAfter = [SELECT Id FROM OrigemProduto__c];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0,lstOrigemBefore.size());
        System.assertEquals(0,lstLogBefore.size());
        System.assertEquals(9, lstOrigemAfter.size());
		System.assertEquals(3,lstLogAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getOrigem400(){
        test.startTest();
        B2BAgrosysOrigemProduto.FiltroRequisicao filtro = new B2BAgrosysOrigemProduto.FiltroRequisicao();
        B2BAgrosysOrigemProduto b2bGetOrigem = new B2BAgrosysOrigemProduto();
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<OrigemProduto__c> lstOrigemBefore = [SELECT Id FROM OrigemProduto__c];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bGetOrigem.getOrigem(filtro);
		List<OrigemProduto__c> lstOrigemAfter = [SELECT Id FROM OrigemProduto__c];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0,lstOrigemBefore.size());
        System.assertEquals(0,lstLogBefore.size());
        System.assertEquals(0, lstOrigemAfter.size());
		System.assertEquals(1,lstLogAfter.size());
        test.stopTest();
    }
}