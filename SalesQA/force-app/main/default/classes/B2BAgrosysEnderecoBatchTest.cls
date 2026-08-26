@isTest
public class B2BAgrosysEnderecoBatchTest {
	
    @testSetup
    static void makeData() {

        List<Account> lstAccountUpsert = new List<Account>();
        for(Integer i = 1; i < 4; i++) {
        	Account newAcc = new Account(
                Name				= 'isTest ' + i,
                CodigoClienteERP__c	= String.valueOf(i),
                ExternalId__c		= String.valueOf(i),
				Empresa__c 			= 1,
				unidade__c 			= 1
            );
            lstAccountUpsert.add(newAcc);
        }
        upsert lstAccountUpsert;

        Cidade__c newCity = new Cidade__c(
            Name = 'isTest',
            UF__c = 'IT',
            CodigoCidade__c = '1',
            CodigoCidadeIBGE__c = 'isTest',
            CodigoPaisSPED__c = 'isTest'
        );
        insert newCity;
    }
    
    @isTest
    static void getEnderecoBatchSuccess01() {
        Account account = [
            SELECT Id, CodigoClienteERP__c
            FROM Account
            WHERE CodigoClienteERP__c = '1'
        ];

        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_endereco":[{"emp_empresa":1,"uni_unidade":1,"cli_codigo":1,"ced_tpender":1,"ced_seq":1,"ced_endereco":"RODBR-040","ced_numero":"","ced_cep":"33820152","cli_bairro":"VENEZA","cid_cidade":1,"end_latitude":"-19.936097255759893","end_longitude":"-43.93439484999999","end_padrao":true,"end_descricao":"","ced_observ":""}]}]}]}');
        mock.setStatusCode(200);
        Test.setMock(HttpCalloutMock.class, mock);
        
		B2BAgrosysEnderecoBatch b2bBatch = new B2BAgrosysEnderecoBatch(account.CodigoClienteERP__c);
        
        Test.startTest();
        Database.executeBatch(b2bBatch);
        Test.stopTest();

        List<ContactPointAddress> lstCPA = [SELECT Id FROM ContactPointAddress];
        System.assertEquals(1, lstCPA.size());

        List<IntegrationLog__c> lstLog = [SELECT Id FROM IntegrationLog__c];
        System.assertEquals(1, lstLog.size());
    }
    
    @isTest
    static void getEnderecoBatchSuccess02() {
        Set<String> setCustomerCode = new Set<String>();
        for(Account acc : [SELECT Id, CodigoClienteERP__c FROM Account LIMIT 1]) {
            setCustomerCode.add(acc.CodigoClienteERP__c);
        }

        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_endereco":[{"emp_empresa":1,"uni_unidade":1,"cli_codigo":1,"ced_tpender":1,"ced_seq":1,"ced_endereco":"isTest","ced_numero":"","ced_cep":"isTest","cli_bairro":"isTest","cid_cidade":1,"end_latitude":"-00.000000000000000","end_longitude":"-00.000000000000000","end_padrao":true,"end_descricao":"","ced_observ":""}]}]}]}');
        mock.setStatusCode(200);
        Test.setMock(HttpCalloutMock.class, mock);
        
		B2BAgrosysEnderecoBatch b2bBatch = new B2BAgrosysEnderecoBatch(setCustomerCode);
        
        Test.startTest();
        Database.executeBatch(b2bBatch);
        Test.stopTest();

        List<ContactPointAddress> lstCPA = [SELECT Id FROM ContactPointAddress];
        System.assertEquals(1, lstCPA.size());

        List<IntegrationLog__c> lstLog = [SELECT Id FROM IntegrationLog__c];
        System.assertEquals(1, lstLog.size());
    }

    @isTest
    static void getEnderecoBatchError01() {
        Account account = [
            SELECT Id, CodigoClienteERP__c
            FROM Account
            WHERE CodigoClienteERP__c = '1'
        ];

        delete [SELECT Id FROM Cidade__c];

        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_endereco":[{"emp_empresa":1,"uni_unidade":1,"cli_codigo":1,"ced_tpender":1,"ced_seq":1,"ced_endereco":"RODBR-040","ced_numero":"","ced_cep":"33820152","cli_bairro":"VENEZA","cid_cidade":1,"end_latitude":"-19.936097255759893","end_longitude":"-43.93439484999999","end_padrao":true,"end_descricao":"","ced_observ":""}]}]}]}');
        mock.setStatusCode(200);
        Test.setMock(HttpCalloutMock.class, mock);
        
        B2BAgrosysEnderecoBatch.forceException = true;
		B2BAgrosysEnderecoBatch b2bBatch = new B2BAgrosysEnderecoBatch(account.CodigoClienteERP__c);
        
        Test.startTest();
        Database.executeBatch(b2bBatch);
        Test.stopTest();
        
        List<ContactPointAddress> lstCPA = [SELECT Id FROM ContactPointAddress];
        System.assertEquals(0, lstCPA.size());

        List<IntegrationLog__c> lstLog = [SELECT Id FROM IntegrationLog__c];
        System.assertEquals(2, lstLog.size());
    }
	
    @isTest
    static void getEnderecoBatchError02() {
        Account account = [
            SELECT Id, Name, CodigoClienteERP__c
            FROM Account
            WHERE CodigoClienteERP__c = '1'
        ];
        account.Name = 'error02';
        update account;

        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"tp_retorno":[{"msg_codigo":1,"msg_descricao":"Integracaorealizadacomsucesso","msg_retorno":[{"tp_endereco":[{"emp_empresa":1,"uni_unidade":1,"cli_codigo":1,"ced_tpender":1,"ced_seq":1,"ced_endereco":"RODBR-040","ced_numero":"","ced_cep":"33820152","cli_bairro":"VENEZA","cid_cidade":1,"end_latitude":"-19.936097255759893","end_longitude":"-43.93439484999999","end_padrao":true,"end_descricao":"","ced_observ":""}]}]}]}');
        mock.setStatusCode(200);
        Test.setMock(HttpCalloutMock.class, mock);
        
        B2BAgrosysEnderecoBatch.forceException = true;
		B2BAgrosysEnderecoBatch b2bBatch = new B2BAgrosysEnderecoBatch(account.CodigoClienteERP__c);
        
        Boolean occurredError = false;
        Test.startTest();
        try {
            Database.executeBatch(b2bBatch);
        } catch (Exception ex) {
            occurredError = true;
        }
        Test.stopTest();

        System.assertEquals(false, occurredError);
    }

}