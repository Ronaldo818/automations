@IsTest
public class B2BAgrosysProdutoTest {
    
    @TestSetup
    static void makeData() {
        Product2 newProduct01 = new Product2(
			Name = 'isTest');
		insert newProduct01;
        
    	Product2 newProduct02 = new Product2(
		Name = 'isTest');
		insert newProduct02;
        
        PricebookEntry standardPrice = new PricebookEntry(
            Pricebook2Id = Test.getStandardPricebookId(),
            Product2Id = newProduct02.Id,
            UnitPrice = 0,
            IsActive = true);
        insert standardPrice;
        
		Pricebook2 newPricebook = new Pricebook2(
			Name = 'isTest');
		insert newPricebook;
      
		PricebookEntry newPricebokEntry = new PricebookEntry(
			UnitPrice = 1,
    		Pricebook2Id = newPricebook.Id,
			Product2Id = newProduct02.Id);
		insert newPricebokEntry;
    }
    
	@isTest
    static void getProduto200(){
        test.startTest();
        B2BAgrosysProduto.FiltroRequisicao filtro = new B2BAgrosysProduto.FiltroRequisicao();
        B2BAgrosysProduto b2bProduto = new B2BAgrosysProduto();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_produto":[{"pro_codigo":1,"pro_descricao":"**FRANGOCONGELADOCXPAPELAO","emp_empresa":1,"uni_unidade":1,"med_codigo":"KG","pro_psunitario":20.0,"pro_unifat":"KG","pro_uniped":"CX","pro_situa":false,"fam_codigo":10,"gru_codigo":98,"lin_codigo":1,"ori_codigo":1,"emb_peso":0,"emb_altura":0,"emb_largura":0,"emb_comprimento":0},{"pro_codigo":1,"pro_descricao":"**FRANGOCONGELADOCXPAPELAO","emp_empresa":1,"uni_unidade":5,"med_codigo":"KG","pro_psunitario":20.0,"pro_unifat":"KG","pro_uniped":"CX","pro_situa":false,"fam_codigo":10,"gru_codigo":98,"lin_codigo":0,"ori_codigo":1,"emb_peso":0,"emb_altura":0,"emb_largura":0,"emb_comprimento":0}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<PricebookEntry> lstPriceEntryBefore = [SELECT Id FROM PricebookEntry];
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bProduto.getProduto(filtro);
		List<PricebookEntry> lstPriceEntryAfter = [SELECT Id FROM PricebookEntry];
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(2, lstPriceEntryBefore.size());
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(4, lstLogAfter.size());
        System.assertEquals(3, lstPriceEntryAfter.size());
        test.stopTest();
    }
    
    @isTest
    static void getProduto400(){
        test.startTest();
        B2BAgrosysProduto.FiltroRequisicao filtro = new B2BAgrosysProduto.FiltroRequisicao();
        B2BAgrosysProduto b2bProduto = new B2BAgrosysProduto();
        
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bProduto.getProduto(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0, lstLogBefore.size());
		System.assertEquals(1, lstLogAfter.size());
        test.stopTest();
    }
}