@IsTest
public class B2BAgrosysEnderecoTest {
    
    @TestSetup
    static void makeData() {
        List<Account> lstAccount = new List<Account>();
        
        for(Integer i = 0; i < 3; i++) {
        	Account acc = new Account(
                Name = 'Customer ' + i,
                CodigoClienteERP__c = String.valueOf(i),
            	Empresa__c = 1,
                Unidade__c = 1);
            lstAccount.add(acc);
        }
        upsert lstAccount;
    }
   
    @isTest
    static void getEndereco200(){
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_endereco":[{"emp_empresa":1,"uni_unidade":1,"cli_codigo":1680,"ced_tpender":1,"ced_seq":1,"ced_endereco":"RODBR-040","ced_numero":"","ced_cep":"33820152","cli_bairro":"VENEZA","cid_cidade":647,"end_latitude":"-19.936097255759893","end_longitude":"-43.93439484999999","end_padrao":true,"end_descricao":"","ced_observ":""},{"emp_empresa":1,"uni_unidade":1,"cli_codigo":1680,"ced_tpender":1,"ced_seq":3,"ced_endereco":"RPRESTANCREDODEALMEIDANEVES","ced_numero":"1041","ced_cep":"35.557000","cli_bairro":"BONFIM","cid_cidade":3463,"end_latitude":"-19.936097255759893","end_longitude":"-43.93439484999999","end_padrao":false,"end_descricao":"","ced_observ":""},{"emp_empresa":1,"uni_unidade":1,"cli_codigo":1680,"ced_tpender":1,"ced_seq":4,"ced_endereco":"RESPLANADA","ced_numero":"","ced_cep":"33.805660","cli_bairro":"VILAESPLANADA","cid_cidade":647,"end_latitude":"-19.936097255759893","end_longitude":"-43.93439484999999","end_padrao":false,"end_descricao":"","ced_observ":""}]}]}]}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);
        
		List<B2BAgrosysRetorno.TpEndereco> b2bRetorno = new List<B2BAgrosysRetorno.TpEndereco>();
        B2BAgrosysEndereco b2BEndereco = new B2BAgrosysEndereco();
        
        test.startTest();
		b2bRetorno = b2BEndereco.getEnderecoCliente(606);
        test.stopTest();
    }
    
    @isTest
    static void getEndereco400(){
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('');
        mock.setStatusCode(400);
        test.setMock(HttpCalloutMock.class, mock);
        
        List<B2BAgrosysRetorno.TpEndereco> b2bRetorno = new List<B2BAgrosysRetorno.TpEndereco>();
        B2BAgrosysEndereco b2BEndereco = new B2BAgrosysEndereco();
        
        test.startTest();
		b2bRetorno = b2BEndereco.getEnderecoCliente(606);
        test.stopTest();
    }
    
}