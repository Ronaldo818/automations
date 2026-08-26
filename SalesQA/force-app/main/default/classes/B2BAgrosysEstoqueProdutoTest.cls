@ISTest
public class B2BAgrosysEstoqueProdutoTest  {
    
	@isTest
    static void getEstoque200(){
        test.startTest();
        B2BAgrosysEstoqueProduto.FiltroRequisicao filtro = new B2BAgrosysEstoqueProduto.FiltroRequisicao();
        B2BAgrosysEstoqueProduto b2bgetEstoque = new B2BAgrosysEstoqueProduto();
        B2BAgrosysRetorno b2bRetorno = new B2BAgrosysRetorno();

		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"errocleber","msg_retorno":[{"tp_estoque":[{"emp_empresa":1,"uni_unidade":0,"pro_codigo":104,"est_dtdisp":"2023-06-01","est_camara":4320.0},{"emp_empresa":1,"uni_unidade":0,"pro_codigo":110,"est_dtdisp":"2023-06-01","est_camara":14160.0},{"emp_empresa":1,"uni_unidade":0,"pro_codigo":165,"est_dtdisp":"2023-06-01","est_camara":0.0}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2bgetEstoque.getEstoque(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0,lstLogBefore.size());
		System.assertEquals(1,lstLogAfter.size());
		System.assertEquals(true,b2bRetorno != null);
        test.stopTest();
    }
    
    @isTest
    static void getEstoque400(){
        test.startTest();
        B2BAgrosysEstoqueProduto.FiltroRequisicao filtro = new B2BAgrosysEstoqueProduto.FiltroRequisicao();
		B2BAgrosysRetorno b2bRetorno = new B2BAgrosysRetorno();
        B2BAgrosysEstoqueProduto b2bgetEstoque = new B2BAgrosysEstoqueProduto();
		HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<IntegrationLog__c> lstLogBefore = [SELECT Id FROM IntegrationLog__c];
		b2bRetorno = b2bgetEstoque.getEstoque(filtro);
		List<IntegrationLog__c> lstLogAfter = [SELECT Id FROM IntegrationLog__c];
        
        System.assertEquals(0,lstLogBefore.size());
		System.assertEquals(1,lstLogAfter.size());
		System.assertEquals(false,b2bRetorno != null);
        test.stopTest();
    }
}