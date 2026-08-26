@isTest
global with sharing class IntegracaoReceitaWSTest {
    private static final String CNPJ_ALFA_MASCARADO = 'AB.12C.D34/0001-81';

    @testSetup
    static void testSetup(){

        Account acc = new Account(

            CNPJ__c = '19248604000180',
            Name = 'Test Acc 01'

        );
        insert acc;

        Lead lead = new Lead(

            LastName = 'Test Lead 01',
            Company = 'Test Company',
            Status = 'Novo'

        );
        insert lead;




    }

    @isTest 
    static void consultarDadosReceitaWSTest() {

        Account acc = [SELECT CNPJ__c FROM Account LIMIT 1];

        Test.setMock(HttpCalloutMock.class, new IntegracaoReceitaWSCalloutMock());
        Test.startTest();
         IntegracaoReceitaWS.consultarDadosReceitaWS('68.940.748/0001-08');
        Test.stopTest();

    }

    @isTest
    static void consultarDadosReceitaWSAlfanumericoTest() {
        Test.setMock(HttpCalloutMock.class, new IntegracaoReceitaWSCalloutAlfanumericoMock());

        Test.startTest();
        IntegracaoReceitaWS.Wrapper result = IntegracaoReceitaWS.consultarDadosReceitaWS(CNPJ_ALFA_MASCARADO);
        Test.stopTest();

        System.assertNotEquals(null, result, 'O retorno da integracao nao deve ser nulo.');
        System.assertEquals(null, result.erro, 'Nao deve retornar erro para o mock alfanumerico.');
        System.assertNotEquals(null, result.dadosGeral, 'Os dados gerais devem ser preenchidos.');
        System.assertEquals('AB12CD34000181', result.dadosGeral.cnpj, 'O endpoint deve receber o CNPJ alfanumerico sem mascara.');
    }

    @isTest
    static void atualizarObjetoTest() {

        Account acc = [SELECT Id FROM Account LIMIT 1];
        Lead lead   = [SELECT Id FROM Lead    LIMIT 1];

        Test.startTest();
        String  recordId2=acc.Id; 
        IntegracaoReceitaWS.DadosReceitaWS registro = new IntegracaoReceitaWS.DadosReceitaWS(); 
        registro.cnpj = '00.000.000/0000-00';
        registro.ultima_atualizacao = '2019-12-19T00:00:00.000+0000';
        registro.numero = '550';
        String dadosQuadSocietario; 
        String atvEconomicaPrincipal; 
        String atvEconomicaSecundaria;


        IntegracaoReceitaWS.atualizarObjeto(recordId2, registro, dadosQuadSocietario, atvEconomicaPrincipal, atvEconomicaSecundaria);
        Test.stopTest();


    }

    @isTest 
    static void consultarDadosObjetoTest() {

        Account acc = [SELECT Id FROM Account LIMIT 1];
        Lead lead   = [SELECT Id FROM Lead    LIMIT 1];

        Test.startTest();
         IntegracaoReceitaWS.consultarDadosObjeto(acc.Id);
        Test.stopTest();

    }

    global class IntegracaoReceitaWSCalloutMock  implements HttpCalloutMock {
    
        global HttpResponse respond(HttpRequest req) {
             
            HttpResponse res = new HttpResponse();
    
            res.setHeader('Content-Type', 'application/json'); 
            res.setBody('{"atividade_principal":[{"text":"Gestão e administração da propriedade imobiliária","code":"68.22-6-00"}],"data_situacao":"13/11/2013","complemento":"SLJ 201","tipo":"MATRIZ","nome":"ADMINISTRADORA NORTH SHOPPING JOQUEI LTDA","uf":"CE","telefone":"(85) 8819-2856","atividades_secundarias":[{"text":"Estacionamento de veículos","code":"52.23-1-00"}],"qsa":[{"qual":"05-Administrador","nome":"MARCELO BAPTISTA CARVALHO"},{"qual":"22-Sócio","qual_rep_legal":"05-Administrador","nome_rep_legal":"MARCOS BAPTISTA CARVALHO","nome":"SPE FORTALEZA SHOPPING SA"},{"qual":"22-Sócio","qual_rep_legal":"05-Administrador","nome_rep_legal":"RODRIGO MORAES MARTINS","nome":"SPE ANDRIOS EMPREENDIMENTOS IMOBILIARIOS S.A."},{"qual":"05-Administrador","nome":"MARCOS BAPTISTA CARVALHO"}],"situacao":"ATIVA","bairro":"JOQUEI CLUBE","logradouro":"AV LINEU MACHADO","numero":"550","cep":"60.520-101","municipio":"FORTALEZA","porte":"DEMAIS","abertura":"13/11/2013","natureza_juridica":"206-2 - Sociedade Empresária Limitada","fantasia":"ADMINISTRADORA NORTH SHOPPING JOQUEI LTDA","cnpj":"19.248.604/0001-80","ultima_atualizacao":"2022-02-16T09:49:14.907Z","status":"OK","email":"","efr":"","motivo_situacao":"","situacao_especial":"","data_situacao_especial":"","capital_social":"1000.00","extra":{},"billing":{"free":true,"database":true}}');
            res.setStatusCode(200);
            return res;
        }
    }

    global class IntegracaoReceitaWSCalloutAlfanumericoMock implements HttpCalloutMock {

        global HttpResponse respond(HttpRequest req) {
            System.assert(req.getEndpoint().endsWith('/AB12CD34000181'),
                'O endpoint deve manter o CNPJ alfanumerico sem mascara.');

            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');
            res.setBody('{"atividade_principal":[{"text":"Teste","code":"01.11-3-01"}],"data_situacao":"13/11/2013","complemento":"","tipo":"MATRIZ","nome":"EMPRESA TESTE","uf":"SP","telefone":"(11) 9999-9999","atividades_secundarias":[],"qsa":[],"situacao":"ATIVA","bairro":"CENTRO","logradouro":"RUA TESTE","numero":"100","cep":"01.010-000","municipio":"SAO PAULO","porte":"DEMAIS","abertura":"13/11/2013","natureza_juridica":"206-2 - Sociedade Empresaria Limitada","fantasia":"EMPRESA TESTE","cnpj":"AB12CD34000181","ultima_atualizacao":"2022-02-16T09:49:14.907Z","status":"OK","email":"","efr":"","motivo_situacao":"","situacao_especial":"","data_situacao_especial":"","capital_social":"1000.00","extra":{},"billing":{"free":true,"database":true}}');
            res.setStatusCode(200);
            return res;
        }
    }

}