@IsTest
public class B2BAgrosysRetornoTest {

	@TestSetup
    static void makeData() {
        
        Pricebook2 newPriceBook = new Pricebook2(Name = 'isTest');
		insert newPriceBook;
        
        CondicaoPagamento__c newCondPagto = new CondicaoPagamento__c(
        	AVista__c = true,
        	CodigoCondicaoPagamento__c = '1',
        	Descricao__c = 'isTest',
        	Name = 'isTest',
        	NumeroParcelas__c = 1);
        insert newCondPagto;
        
        Account newAccount = new Account(
        	Name = 'isTest',
        	Unidade__c = 1,
        	Empresa__c = 1,
        	CodigoClienteERP__c = '1',
        	CondicaoPagamento__c = newCondPagto.Id,
        	Ocorrencia__c = 1,
        	Lista__c = 27);
        insert newAccount;
        
        Order newOrder = new Order(
        	EffectiveDate = Date.today(),
        	Status = 'Draft',
        	ValorCashback__c = 2,
        	AccountId = newAccount.Id,
        	Pricebook2Id = newPriceBook.Id);
        insert newOrder;
    }
    
    @isTest
    static void setPedidoSuccess01(){
        B2BAgrosys.HttpParameters parameters = new B2BAgrosys.HttpParameters();
        String json = '{"tp_retorno":[{"msg_codigo" : 1,"msg_descricao":"Integracao realizada com sucesso","msg_retorno":[{"tp_clineg": [{"temnegociacao": false}]}]}]}';
        B2BAgrosysRetorno retorno = new B2BAgrosysRetorno();
        retorno = B2BAgrosysRetorno.parse(json, parameters);
    }
    
    @isTest
    static void setPedidoError01(){
        B2BAgrosys.HttpParameters parameters = new B2BAgrosys.HttpParameters();
        String json = 'error';
        B2BAgrosysRetorno retorno = new B2BAgrosysRetorno();
        retorno = B2BAgrosysRetorno.parse(json, parameters);
    }
    
    @isTest
    static void setPedidoError02(){
        IntegrationLog__c log = new IntegrationLog__c();
        insert log;
        
        B2BAgrosys.HttpParameters parameters = new B2BAgrosys.HttpParameters();
        parameters.log = log;
        
        String json = 'error';
        B2BAgrosysRetorno retorno = new B2BAgrosysRetorno();
        retorno = B2BAgrosysRetorno.parse(json, parameters);
    }
}